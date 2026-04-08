let updateInterval;
const API_URL_1 = "https://script.google.com/macros/s/AKfycbxULPYNOZv_ZrY81RL-YJ14LJcd5MMPUPU6Esc7eQer9w3s16m1hh7h6jUk13_zy1Y/exec";
const API_URL_2 = "https://script.google.com/macros/s/AKfycbzEDVFJ7mwPB5acuZWo4uuphMGMP3YyRzLw3gIM61T6Y0tvngkAKdpcD_7Tyz1Nj8dO/exec";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Buton Tıklamaları
        document.getElementById("btn-1").onclick = () => fetchData(API_URL_1, "GoogleVerileri");
        document.getElementById("btn-2").onclick = () => fetchData(API_URL_2, "GoogleVerileri2");

        // Zamanlayıcı Ayarları
        document.getElementById("auto-update-cb").onchange = handleTimer;
        document.getElementById("interval-select").onchange = handleTimer;
    }
});

async function fetchData(url, sheetName) {
    const status = document.getElementById("status");
    try {
        status.innerText = `Durum: ${sheetName} senkronize ediliyor...`;
        status.style.color = "blue";

        const response = await fetch(url);
        const data = await response.json();

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            let sheet = sheets.getItemOrNullObject(sheetName);
            await context.sync();

            if (sheet.isNullObject) {
                sheet = sheets.add(sheetName);
            }

            sheet.getUsedRange().clear(Excel.ClearApplyTo.contents);
            
            const rowCount = data.length;
            const colCount = data[0].length;
            const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
            
            range.values = data;
            range.format.autofitColumns();
            
            await context.sync();
        });

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()}`;
        status.style.color = "green";
    } catch (error) {
        console.error(error);
        status.innerText = "Hata: Veri çekilemedi!";
        status.style.color = "red";
    }
}

function handleTimer() {
    const isChecked = document.getElementById("auto-update-cb").checked;
    const intervalMs = parseInt(document.getElementById("interval-select").value);
    const timerText = document.getElementById("timer-text");

    clearInterval(updateInterval);

    if (isChecked) {
        const mins = intervalMs / 60000;
        timerText.innerText = `Aktif: ${mins} dk'da bir yenileniyor...`;
        timerText.style.color = "green";
        
        // İlk veriyi hemen çek, sonra periyoda başla
        fetchData(API_URL_1, "GoogleVerileri");
        
        updateInterval = setInterval(() => {
            fetchData(API_URL_1, "GoogleVerileri");
        }, intervalMs);
    } else {
        timerText.innerText = "Zamanlayıcı Kapalı";
        timerText.style.color = "#666";
    }
}
