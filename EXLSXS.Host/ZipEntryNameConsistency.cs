using System.Buffers.Binary;
using System.Text;

namespace EXLSXS.Host;

// 更新パッケージ (zip) の構造を自前で読み、.NET の ZipArchive (署名を検査する側) と
// Velopack の Update.exe (Rust の zip クレート、実際に展開する側) が同じアーカイブ・同じ名前を
// 読むことを確かめる。食い違いがあると「検査側には非 PE に見えて署名検査を素通りし、展開側では
// coreclr.dll などとして書き出される」細工ができてしまう。ZipArchiveEntry は拡張フィールドや
// EOCD の位置を公開しないため、署名検査の前にここで直接確かめる。
//
// 拒否するもの (どれも正規の vpk 出力には現れない):
//   - 末尾ちょうど 22 バイト以外の EOCD (コメント付き)。zip クレートは末尾から探すので、コメント内の EOCD を先に試す
//   - Zip64 (EOCD・中央ディレクトリの上限値)。zip クレートは size の上限値でも Zip64 を探すなど、判定が .NET と異なる
//   - 採用した EOCD より手前にある、展開側の戻り先になりうる別の EOCD。zip クレートは中央ディレクトリの
//     読み取りに失敗すると手前の EOCD を探し直し、記載 offset が外れていても前方を探して中央ヘッダーを使う
//   - 中央ディレクトリ・ローカルヘッダーの Info-ZIP Unicode Path 拡張フィールド (0x7075) と壊れた拡張フィールド。
//     zip クレートはこの値でヘッダーの名前を置き換えるが、.NET は無視する
//   - 中央ディレクトリの名前の並びと .NET の Entries の不一致、ローカルヘッダーとの名前の不一致
internal static class ZipEntryNameConsistency
{
    private const uint LocalHeaderSignature = 0x04034b50;
    private const uint CentralHeaderSignature = 0x02014b50;
    private const uint EndOfCentralDirectorySignature = 0x06054b50;
    private const uint Zip64LocatorSignature = 0x07064b50;
    private const ushort UnicodePathExtraFieldId = 0x7075;
    private const int EndOfCentralDirectorySize = 22;
    private const int Zip64LocatorSize = 20;
    private const int CentralHeaderFixedSize = 46;
    private const int LocalHeaderFixedSize = 30;

    public static bool Verify(string packagePath, IReadOnlyList<string> dotnetEntryNames, out string message)
    {
        try
        {
            using var stream = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return VerifyCore(stream, dotnetEntryNames, out message);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or ArgumentOutOfRangeException)
        {
            message = $"Update package zip structure could not be verified: {ex.Message}";
            return false;
        }
    }

    private static bool VerifyCore(FileStream stream, IReadOnlyList<string> dotnetEntryNames, out string message)
    {
        var length = stream.Length;
        if (length < EndOfCentralDirectorySize)
        {
            message = "Update package is too small to be a zip archive.";
            return false;
        }

        // 1. EOCD はファイル末尾ちょうど 22 バイト (コメント長 0) だけを認める。
        var eocdPosition = length - EndOfCentralDirectorySize;
        var eocd = ReadExactly(stream, eocdPosition, EndOfCentralDirectorySize);
        if (BinaryPrimitives.ReadUInt32LittleEndian(eocd) != EndOfCentralDirectorySignature
            || BinaryPrimitives.ReadUInt16LittleEndian(eocd.AsSpan(20)) != 0)
        {
            message = "Update package must end with an end of central directory record without a comment.";
            return false;
        }

        // 2. Zip64 は使わない (分割アーカイブも同様に拒否する)。
        var diskNumber = BinaryPrimitives.ReadUInt16LittleEndian(eocd.AsSpan(4));
        var centralDirectoryDisk = BinaryPrimitives.ReadUInt16LittleEndian(eocd.AsSpan(6));
        var entriesOnDisk = BinaryPrimitives.ReadUInt16LittleEndian(eocd.AsSpan(8));
        var entryCount = BinaryPrimitives.ReadUInt16LittleEndian(eocd.AsSpan(10));
        var centralDirectorySize = BinaryPrimitives.ReadUInt32LittleEndian(eocd.AsSpan(12));
        var centralDirectoryOffset = BinaryPrimitives.ReadUInt32LittleEndian(eocd.AsSpan(16));
        if (UsesZip64(entryCount, centralDirectorySize, centralDirectoryOffset)
            || diskNumber != 0 || centralDirectoryDisk != 0 || entriesOnDisk != entryCount)
        {
            message = "Update package must be a single-disk zip archive without Zip64.";
            return false;
        }

        if (entryCount != dotnetEntryNames.Count)
        {
            message = $"Update package central directory has {entryCount} entries, but {dotnetEntryNames.Count} were read.";
            return false;
        }

        // 3. 手前に展開側の戻り先になりうる別の EOCD が無いこと。
        if (TryFindAlternateEndOfCentralDirectory(stream, eocdPosition, out var alternatePosition))
        {
            message = $"Update package contains another end of central directory record at offset {alternatePosition}, which an updater could fall back to.";
            return false;
        }

        // 4. 中央ディレクトリと各ローカルヘッダーの名前・拡張フィールドを確かめる。
        long position = centralDirectoryOffset;
        for (var index = 0; index < dotnetEntryNames.Count; index++)
        {
            var header = ReadExactly(stream, position, CentralHeaderFixedSize);
            if (BinaryPrimitives.ReadUInt32LittleEndian(header) != CentralHeaderSignature)
            {
                message = "Update package central directory entry is malformed.";
                return false;
            }

            var compressedSize = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(20));
            var uncompressedSize = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(24));
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(28));
            var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(30));
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(32));
            var localHeaderOffset = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(42));

            var nameBytes = ReadExactly(stream, position + CentralHeaderFixedSize, nameLength);
            var extra = ReadExactly(stream, position + CentralHeaderFixedSize + nameLength, extraLength);
            var name = Encoding.UTF8.GetString(nameBytes);

            if (!string.Equals(name, dotnetEntryNames[index], StringComparison.Ordinal))
            {
                message = $"Update package entry #{index} is read as '{dotnetEntryNames[index]}' by .NET but '{name}' in the central directory.";
                return false;
            }

            if (compressedSize == uint.MaxValue || uncompressedSize == uint.MaxValue || localHeaderOffset == uint.MaxValue)
            {
                message = $"Update package entry '{name}' uses Zip64 sizes or offsets.";
                return false;
            }

            if (ContainsExtraField(extra, UnicodePathExtraFieldId))
            {
                message = $"Update package entry '{name}' has an Info-ZIP Unicode Path field, which the updater would use instead of the verified name.";
                return false;
            }

            if (!VerifyLocalHeader(stream, localHeaderOffset, nameBytes, name, out message))
            {
                return false;
            }

            position += CentralHeaderFixedSize + nameLength + extraLength + commentLength;
        }

        message = "Update package entry names are consistent.";
        return true;
    }

    private static bool UsesZip64(ushort entryCount, uint centralDirectorySize, uint centralDirectoryOffset)
        => entryCount == ushort.MaxValue || centralDirectorySize == uint.MaxValue || centralDirectoryOffset == uint.MaxValue;

    private static bool VerifyLocalHeader(FileStream stream, long offset, byte[] centralNameBytes, string name, out string message)
    {
        var header = ReadExactly(stream, offset, LocalHeaderFixedSize);
        if (BinaryPrimitives.ReadUInt32LittleEndian(header) != LocalHeaderSignature)
        {
            message = $"Update package entry '{name}' points to a malformed local header.";
            return false;
        }

        var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(26));
        var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(28));
        var localNameBytes = ReadExactly(stream, offset + LocalHeaderFixedSize, nameLength);
        if (!localNameBytes.AsSpan().SequenceEqual(centralNameBytes))
        {
            message = $"Update package entry '{name}' has a different name in its local header.";
            return false;
        }

        var extra = ReadExactly(stream, offset + LocalHeaderFixedSize + nameLength, extraLength);
        if (ContainsExtraField(extra, UnicodePathExtraFieldId))
        {
            message = $"Update package entry '{name}' has an Info-ZIP Unicode Path field in its local header.";
            return false;
        }

        message = "";
        return true;
    }

    // 採用した EOCD より手前の EOCD シグネチャを、zip クレートが戻り先として使いうる条件より広く拒否する。
    //   - エントリ数 1 以上で記載 offset が自分より手前、かつ offset から自分までのどこかに中央ヘッダーのシグネチャがある
    //     (zip クレートは記載 offset が外れても前方を探して最初の中央ヘッダーを使う。size は位置決めに使わない)
    //   - Zip64 を示す上限値を持ち、20 バイト手前に Zip64 locator がある (locator 経由で別の中央ディレクトリを指せる)
    // 圧縮データに偶然 4 バイトが一致しても、さらにその手前に中央ヘッダーのシグネチャが要るので正規パッケージでは起きにくい
    // (正規の中央ディレクトリは EOCD の直前にあり、手前の偶然一致より後ろに位置する)。
    private static bool TryFindAlternateEndOfCentralDirectory(FileStream stream, long acceptedEocdPosition, out long candidatePosition)
    {
        candidatePosition = -1;
        var centralSignaturePositions = new List<long>();
        var eocdSignaturePositions = new List<long>();
        ScanSignatures(stream, acceptedEocdPosition, centralSignaturePositions, eocdSignaturePositions);

        foreach (var position in eocdSignaturePositions)
        {
            // 候補は採用した EOCD より手前にあるので、22 バイトは必ずファイル内に収まる。
            var record = ReadExactly(stream, position, EndOfCentralDirectorySize);
            var entriesOnDisk = BinaryPrimitives.ReadUInt16LittleEndian(record.AsSpan(8));
            var entries = BinaryPrimitives.ReadUInt16LittleEndian(record.AsSpan(10));
            var size = BinaryPrimitives.ReadUInt32LittleEndian(record.AsSpan(12));
            var offset = BinaryPrimitives.ReadUInt32LittleEndian(record.AsSpan(16));

            var hasCentralHeaderBefore = entries > 0
                && offset < position
                && HasPositionInRange(centralSignaturePositions, offset, position);
            var hasZip64Locator = UsesZip64(entries, size, offset)
                && position >= Zip64LocatorSize
                && BinaryPrimitives.ReadUInt32LittleEndian(ReadExactly(stream, position - Zip64LocatorSize, 4)) == Zip64LocatorSignature;
            // zip クレートは総数 0 の EOCD を中央ヘッダーを探さずに採用し、その後はディスク上の数だけエントリを読む。
            // 記載 offset の向きに関係なく戻り先になりうるので、この組み合わせは無条件で拒否する。
            var readsEntriesDespiteZeroTotal = entries == 0 && entriesOnDisk > 0;

            if (hasCentralHeaderBefore || hasZip64Locator || readsEntriesDespiteZeroTotal)
            {
                candidatePosition = position;
                return true;
            }
        }

        return false;
    }

    // [0, end) を 1MB ずつ読み、中央ヘッダーと EOCD のシグネチャの位置を集める (境界をまたぐ 3 バイトは持ち越す)。
    private static void ScanSignatures(FileStream stream, long end, List<long> centralPositions, List<long> eocdPositions)
    {
        const int chunkSize = 1 << 20;
        var buffer = new byte[chunkSize + 3];
        long bufferStart = 0;
        var carried = 0;
        stream.Position = 0;
        while (bufferStart + carried < end)
        {
            var toRead = (int)Math.Min(chunkSize, end - bufferStart - carried);
            stream.ReadExactly(buffer.AsSpan(carried, toRead));
            var available = carried + toRead;
            for (var i = 0; i + 4 <= available; i++)
            {
                var signature = BinaryPrimitives.ReadUInt32LittleEndian(buffer.AsSpan(i));
                if (signature == CentralHeaderSignature)
                {
                    centralPositions.Add(bufferStart + i);
                }
                else if (signature == EndOfCentralDirectorySignature)
                {
                    eocdPositions.Add(bufferStart + i);
                }
            }

            carried = Math.Min(3, available);
            Array.Copy(buffer, available - carried, buffer, 0, carried);
            bufferStart += available - carried;
        }
    }

    // sortedPositions (昇順) に start <= p < end の位置があるか。
    private static bool HasPositionInRange(List<long> sortedPositions, long start, long end)
    {
        var index = sortedPositions.BinarySearch(start);
        if (index < 0)
        {
            index = ~index;
        }

        return index < sortedPositions.Count && sortedPositions[index] < end;
    }

    // 壊れた拡張フィールド (宣言サイズが領域を超える) は、その先に 0x7075 を隠せるため「含む」扱いで拒否する。
    private static bool ContainsExtraField(byte[] extra, ushort fieldId)
    {
        var position = 0;
        while (position < extra.Length)
        {
            if (position + 4 > extra.Length)
            {
                return true;
            }

            var id = BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(position));
            var size = BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(position + 2));
            if (position + 4 + size > extra.Length || id == fieldId)
            {
                return true;
            }

            position += 4 + size;
        }

        return false;
    }

    private static byte[] ReadExactly(FileStream stream, long offset, int count)
    {
        if (offset < 0 || offset + count > stream.Length)
        {
            throw new InvalidDataException("Zip structure points outside the file.");
        }

        var buffer = new byte[count];
        stream.Position = offset;
        stream.ReadExactly(buffer);
        return buffer;
    }
}
