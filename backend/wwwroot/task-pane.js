Office.onReady((info) => {
    console.log("Office.onReady triggered, host:", info.host);
    if (info.host === Office.HostType.Excel) {
        console.log("Excel host detected");
        document.getElementById("runButton").onclick = runFinish;
        document.getElementById("adjustFontCheck").onchange = toggleFontSection;
        loadFonts();
    } else {
        console.log("Not Excel host, loading fonts anyway for testing");
        loadFonts();
    }
});

function toggleFontSection() {
    const isChecked = document.getElementById("adjustFontCheck").checked;
    document.getElementById("fontSection").style.display = isChecked ? "block" : "none";
    document.getElementById("fontSelect").disabled = !isChecked;
}

async function loadFonts() {
    try {
        console.log("loadFonts started");
        const response = await fetch('/api/finish/fonts');
        console.log("fetch response:", response.status);
        if (response.ok) {
            const fonts = await response.json();
            console.log("Fonts loaded:", fonts.length);
            const select = document.getElementById("fontSelect");
            select.innerHTML = "";
            fonts.forEach(font => {
                const option = document.createElement("option");
                option.text = font;
                option.value = font;
                select.add(option);
            });
            // デフォルト選択
            if (fonts.length > 0) {
                select.value = fonts[0];
            }
            console.log("Font select populated");
        } else {
            console.error("fetch failed:", response.status);
        }
    } catch (error) {
        console.error("Failed to load fonts", error);
        const select = document.getElementById("fontSelect");
        select.innerHTML = "<option>Failed to load</option>";
    }
}

async function runFinish() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const isAdjustFont = document.getElementById("adjustFontCheck").checked;
            const fontName = document.getElementById("fontSelect").value;
            // const zoom = parseInt(document.getElementById("zoomSelect").value);
            // const view = document.getElementById("viewSelect").value;

            // 逆順に処理（最後を先頭にするため、あるいはVSTOロジックに合わせる）
            // VSTO: count down to 1.
            // ここではすべてのシートを処理して、最後に最初のシートをアクティブにする
            
            for (let sheet of sheets.items) {
                sheet.activate();
                
                // A1選択
                const range = sheet.getRange("A1");
                range.select();

                // フォント変更
                if (isAdjustFont) {
                    sheet.getUsedRange().format.font.name = fontName;
                }

                // JS APIでのView/Zoom変更は制限があるため、現状はスキップ
                // 将来的にAPIがサポートされたら追加
                // console.log(`Processed ${sheet.name}: Zoom=${zoom}, View=${view}`);
            }

            // 最初のシートをアクティブに
            if (sheets.items.length > 0) {
                sheets.items[0].activate();
            }

            await context.sync();
            
            document.getElementById("status").innerText = "完了しました！";
            setTimeout(() => document.getElementById("status").innerText = "", 3000);
        });
    } catch (error) {
        console.error(error);
        document.getElementById("status").innerText = "エラーが発生しました: " + error.message;
    }
}
