using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace EXLSXS.Host;

internal static class AuthenticodeTrustVerifier
{
    private const uint WtdUiNone = 2;
    private const uint WtdRevokeWholeChain = 1;
    private const uint WtdChoiceFile = 1;
    private const uint WtdStateActionVerify = 1;
    private const uint WtdStateActionClose = 2;
    private const uint WtdRevocationCheckChainExcludeRoot = 0x00000080;

    private static readonly Guid WinTrustActionGenericVerifyV2 =
        new("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");

    internal static bool VerifyEmbeddedSignature(string filePath, out string failureReason)
    {
        X509Certificate2? signerCertificate = null;
        try
        {
            return TryVerifyEmbeddedSignature(filePath, out signerCertificate, out failureReason);
        }
        finally
        {
            signerCertificate?.Dispose();
        }
    }

    internal static bool TryVerifyEmbeddedSignature(
        string filePath,
        out X509Certificate2? signerCertificate,
        out string failureReason)
    {
        signerCertificate = null;
        if (!OperatingSystem.IsWindows())
        {
            failureReason = "Authenticode verification is only available on Windows.";
            return false;
        }

        var fileInfo = new WinTrustFileInfo(filePath);
        var fileInfoPointer = Marshal.AllocHGlobal(Marshal.SizeOf<WinTrustFileInfo>());
        try
        {
            Marshal.StructureToPtr(fileInfo, fileInfoPointer, fDeleteOld: false);
            var trustData = new WinTrustData(fileInfoPointer);
            var action = WinTrustActionGenericVerifyV2;
            var result = WinVerifyTrust(new IntPtr(-1), ref action, ref trustData);
            try
            {
                if (result != 0)
                {
                    failureReason = $"{new Win32Exception(result).Message} (0x{result:X8})";
                    return false;
                }

                if (!TryGetVerifiedSignerCertificate(trustData.StateData, out signerCertificate, out failureReason))
                {
                    return false;
                }

                failureReason = "";
                return true;
            }
            finally
            {
                if (trustData.StateData != IntPtr.Zero)
                {
                    trustData.StateAction = WtdStateActionClose;
                    _ = WinVerifyTrust(new IntPtr(-1), ref action, ref trustData);
                }
            }
        }
        finally
        {
            Marshal.DestroyStructure<WinTrustFileInfo>(fileInfoPointer);
            Marshal.FreeHGlobal(fileInfoPointer);
        }
    }

    private static bool TryGetVerifiedSignerCertificate(
        IntPtr stateData,
        out X509Certificate2? signerCertificate,
        out string failureReason)
    {
        signerCertificate = null;
        if (stateData == IntPtr.Zero)
        {
            failureReason = "WinVerifyTrust did not return provider state data.";
            return false;
        }

        var providerData = WTHelperProvDataFromStateData(stateData);
        var providerSigner = providerData == IntPtr.Zero
            ? IntPtr.Zero
            : WTHelperGetProvSignerFromChain(providerData, 0, counterSigner: false, 0);
        var providerCertificatePointer = providerSigner == IntPtr.Zero
            ? IntPtr.Zero
            : WTHelperGetProvCertFromChain(providerSigner, 0);
        if (providerCertificatePointer == IntPtr.Zero)
        {
            failureReason = "WinVerifyTrust did not return the verified signer certificate.";
            return false;
        }

        try
        {
            var providerCertificate = Marshal.PtrToStructure<CryptProviderCertificate>(providerCertificatePointer);
            if (providerCertificate.CertificateContext == IntPtr.Zero)
            {
                failureReason = "WinVerifyTrust returned an empty signer certificate context.";
                return false;
            }

            signerCertificate = new X509Certificate2(providerCertificate.CertificateContext);
            failureReason = "";
            return true;
        }
        catch (CryptographicException ex)
        {
            failureReason = $"The verified signer certificate could not be read: {ex.Message}";
            return false;
        }
    }

    [DllImport("wintrust.dll", ExactSpelling = true, PreserveSig = true)]
    private static extern int WinVerifyTrust(
        IntPtr windowHandle,
        [In] ref Guid actionId,
        ref WinTrustData trustData);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperProvDataFromStateData(IntPtr stateData);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvSignerFromChain(
        IntPtr providerData,
        uint signerIndex,
        [MarshalAs(UnmanagedType.Bool)] bool counterSigner,
        uint counterSignerIndex);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvCertFromChain(IntPtr providerSigner, uint certificateIndex);

    [StructLayout(LayoutKind.Sequential)]
    private struct CryptProviderCertificate
    {
        public uint Size;
        public IntPtr CertificateContext;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WinTrustFileInfo
    {
        public WinTrustFileInfo(string filePath)
        {
            Size = (uint)Marshal.SizeOf<WinTrustFileInfo>();
            FilePath = filePath;
            FileHandle = IntPtr.Zero;
            KnownSubject = IntPtr.Zero;
        }

        public uint Size;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string FilePath;

        public IntPtr FileHandle;
        public IntPtr KnownSubject;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WinTrustData
    {
        public WinTrustData(IntPtr fileInfoPointer)
        {
            Size = (uint)Marshal.SizeOf<WinTrustData>();
            PolicyCallbackData = IntPtr.Zero;
            SipClientData = IntPtr.Zero;
            UiChoice = WtdUiNone;
            RevocationChecks = WtdRevokeWholeChain;
            UnionChoice = WtdChoiceFile;
            FileInfoPointer = fileInfoPointer;
            StateAction = WtdStateActionVerify;
            StateData = IntPtr.Zero;
            UrlReference = IntPtr.Zero;
            ProviderFlags = WtdRevocationCheckChainExcludeRoot;
            UiContext = 0;
            SignatureSettings = IntPtr.Zero;
        }

        public uint Size;
        public IntPtr PolicyCallbackData;
        public IntPtr SipClientData;
        public uint UiChoice;
        public uint RevocationChecks;
        public uint UnionChoice;
        public IntPtr FileInfoPointer;
        public uint StateAction;
        public IntPtr StateData;
        public IntPtr UrlReference;
        public uint ProviderFlags;
        public uint UiContext;
        public IntPtr SignatureSettings;
    }
}
