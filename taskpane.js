let updateInterval1;
let updateInterval2;
let updateInterval3; 

const API_URL_1 = "https://script.google.com/macros/s/AKfycbxULPYNOZv_ZrY81RL-YJ14LJcd5MMPUPU6Esc7eQer9w3s16m1hh7h6jUk13_zy1Y/exec";
const API_URL_2 = "https://script.google.com/macros/s/AKfycbzEDVFJ7mwPB5acuZWo4uuphMGMP3YyRzLw3gIM61T6Y0tvngkAKdpcD_7Tyz1Nj8dO/exec";
const API_URL_3 = "https://script.google.com/macros/s/AKfycbyJ9r2ijTPYn_v_2kxzC-I9mXXgtS8k9hOQcxV7qQ2Mc6eWfSMbdXGmF9yu6dP-BiQD/exec"; 

const SAYFALAR = [
    "KONTROL", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", 
    "P9", "P10", "P11", "P12", "P13", "P14", "P15", "P16", "P17"
];

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-1").onclick = () => fetchData(API_URL_1, "YahooVerileri");
        document.getElementById("btn-2").onclick = () => fetchData(API_URL_2, "FinansVerileri");
        document.getElementById("btn-3").onclick = () => fetchPortfolioData(); 

        document.getElementById("auto-update-cb-1").onchange = () => handleTimer(1);
        document.getElementById("interval-select-1").onchange = () => handleTimer(1);

        document.getElementById("auto-update-cb-2").onchange = () => handleTimer(2);
        document.getElementById("interval-select-2").onchange = () => handleTimer(2);

        document.getElementById("auto-update-cb-3").onchange = () => handleTimer(3);
        document.getElementById("interval-select-3").onchange = () => handleTimer(3);
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

async function fetchPortfolioData() {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: Portföy (18 Sayfa) güncelleniyor...`;
        status.style.color = "#5c2d91"; 

        await Excel.run(async (context) => {
            for (let sayfaAdi of SAYFALAR) {
                let url = `${API_URL_3}?sayfaAdi=${sayfaAdi}`;
                let response = await fetch(url);
                let data = await response.json();

                if (data.error || !data || data.length === 0) continue;

                const sheets = context.workbook.worksheets;
                let sheet = sheets.getItemOrNullObject(sayfaAdi);
                await context.sync();

                if (sheet.isNullObject) {
                    sheet = sheets.add(sayfaAdi);
                }

                sheet.getUsedRange().clear(Excel.ClearApplyTo.contents);
                const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
                range.values = data;
                range.format.autofitColumns();
            }
            await context.sync();
        });

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (Portföy)`;
        status.style.color = "#107c41"; 
    } catch (error) {
        status.innerText = "Hata: Portföy bağlantısı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

function handleTimer(id) {
    const isChecked = document.getElementById(`auto-update-cb-${id}`).checked;
    const intervalMs = parseInt(document.getElementById(`interval-select-${id}`).value);
    const timerText = document.getElementById(`timer-text-${id}`);

    if (id === 1) clearInterval(updateInterval1);
    else if (id === 2) clearInterval(updateInterval2);
    else if (id === 3) clearInterval(updateInterval3);

    if (isChecked) {
        // Alt kısımdaki bilgilendirme yazısını seçime göre daha düzgün yazdırıyoruz
        let sureMetni = "";
        if (intervalMs === 86400000) {
            sureMetni = "1 Gün";
        } else if (intervalMs === 3600000) {
            sureMetni = "1 Saat";
        } else {
            sureMetni = `${intervalMs / 60000} Dakika`;
        }

        timerText.innerText = `Aktif: ${sureMetni}'da bir yenileniyor.`;
        timerText.style.color = id === 3 ? "#5c2d91" : "#217346"; 
        
        let interval;

        if (id === 3) {
            fetchPortfolioData(); 
            interval = setInterval(() => {
                fetchPortfolioData();
            }, intervalMs);
            updateInterval3 = interval;
        } 
        else {
            const targetUrl = id === 1 ? API_URL_1 : API_URL_2;
            const targetSheet = id === 1 ? "YahooVerileri" : "FinansVerileri";
            
            fetchData(targetUrl, targetSheet); 
            interval = setInterval(() => {
                fetchData(targetUrl, targetSheet);
            }, intervalMs);

            if (id === 1) updateInterval1 = interval;
            else updateInterval2 = interval;
        }

    } else {
        timerText.innerText = "Zamanlayıcı Kapalı";
        timerText.style.color = "#888";
    }
}
