let updateInterval1;
let updateInterval2;

const API_URL_1 = "https://script.google.com/macros/s/AKfycbxULPYNOZv_ZrY81RL-YJ14LJcd5MMPUPU6Esc7eQer9w3s16m1hh7h6jUk13_zy1Y/exec";
const API_URL_2 = "https://script.google.com/macros/s/AKfycbzEDVFJ7mwPB5acuZWo4uuphMGMP3YyRzLw3gIM61T6Y0tvngkAKdpcD_7Tyz1Nj8dO/exec";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-1").onclick = () => fetchData(API_URL_1, "YahooVerileri");
        document.getElementById("btn-2").onclick = () => fetchData(API_URL_2, "FinansVerileri");

        document.getElementById("auto-update-cb-1").onchange = () => handleTimer(1);
        document.getElementById("interval-select-1").onchange = () => handleTimer(1);

        document.getElementById("auto-update-cb-2").onchange = () => handleTimer(2);
        document.getElementById("interval-select-2").onchange = () => handleTimer(2);
    }
});

async function fetchData(url, sheetName) {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: ${sheetName} güncelleniyor...`;
        
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
            const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
            range.values = data;
            range.format.autofitColumns();
            
            await context.sync();
        });

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (${sheetName})`;
        status.style.color = "#217346";
    } catch (error) {
        status.innerText = "Hata: Bağlantı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

function handleTimer(id) {
    const isChecked = document.getElementById(`auto-update-cb-${id}`).checked;
    const intervalMs = parseInt(document.getElementById(`interval-select-${id}`).value);
    const timerText = document.getElementById(`timer-text-${id}`);
    const targetUrl = id === 1 ? API_URL_1 : API_URL_2;
    const targetSheet = id === 1 ? "YahooVerileri" : "FinansVerileri";

    if (id === 1) clearInterval(updateInterval1);
    else clearInterval(updateInterval2);

    if (isChecked) {
        const mins = intervalMs / 60000;
        timerText.innerText = `Aktif: ${mins} dk'da bir yenileniyor.`;
        timerText.style.color = "#217346";
        
        fetchData(targetUrl, targetSheet);
        
        const interval = setInterval(() => {
            fetchData(targetUrl, targetSheet);
        }, intervalMs);

        if (id === 1) updateInterval1 = interval;
        else updateInterval2 = interval;
    } else {
        timerText.innerText = "Zamanlayıcı Kapalı";
        timerText.style.color = "#888";
    }
}
