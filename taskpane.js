let updateInterval1;
let updateInterval2;
let updateInterval3; 
let updateInterval4; 
let updateInterval5; // 5. Kart (Hacim 100 Gün) için eklendi

const API_URL_1 = "https://script.google.com/macros/s/AKfycbxULPYNOZv_ZrY81RL-YJ14LJcd5MMPUPU6Esc7eQer9w3s16m1hh7h6jUk13_zy1Y/exec";
const API_URL_2 = "https://script.google.com/macros/s/AKfycbzEDVFJ7mwPB5acuZWo4uuphMGMP3YyRzLw3gIM61T6Y0tvngkAKdpcD_7Tyz1Nj8dO/exec";
const API_URL_3 = "https://script.google.com/macros/s/AKfycbyJ9r2ijTPYn_v_2kxzC-I9mXXgtS8k9hOQcxV7qQ2Mc6eWfSMbdXGmF9yu6dP-BiQD/exec"; 

// 5. Kart (Hacim 100 Gün) için paylaştığınız Web URL adresi
const API_URL_5 = "https://script.google.com/macros/s/AKfycbz0G3r_f_TvDKS9gHzKWbgE3mIJMk5-M1ACxJgL-udP-9PpINqYVHpIRJSNf3LpYKWD/exec";

// 300 Gün için Sayfalar (P Serisi)
const SAYFALAR = [
    "KONTROL", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", 
    "P9", "P10", "P11", "P12", "P13", "P14", "P15", "P16", "P17"
];

// 1500 Gün için Özel URL'ler
const URLS_1500 = [
    { name: "P1", url: "https://script.google.com/macros/s/AKfycbz8LhJr366qJDHi-79GWd2qUzTgSttOvGFLPs2C3DKIAPkqJtg-gMptnmomNO-Y309g/exec" },
    { name: "P2", url: "https://script.google.com/macros/s/AKfycbwb-ERcOL0NPnz5nCQxDNvOCgIjonAkXt9iTf7p11ozhF-ov7NZDSiPGG3ETlXDEHfi/exec" },
    { name: "P3", url: "https://script.google.com/macros/s/AKfycbxpuGU30vS25m1U_ocyUa90GoOrap-hnRFjqVV1_2x_G_V_xmQLXT-plTEjvexY26_MqQ/exec" },
    { name: "P4", url: "https://script.google.com/macros/s/AKfycbylEcpn5T3OglrnAH6JFkZ-kH2B35COILQFJ_h4NalgI8K75UtM3dFt1TgtcV9u1SLMDg/exec" },
    { name: "P5", url: "https://script.google.com/macros/s/AKfycbxbgndhBDgnAJTNfOlacI0awTBd9lG2Ey4U3VfOzHC0ydx5afv9o9FERmdy9G4gt2NO/exec" }
];

// Hacim 100 Gün için Excel Sayfaları (Görselinizdeki H Serisi + KONTROL)
const HACIM_SAYFALAR = [
    "KONTROL", "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", 
    "H9", "H10", "H11", "H12", "H13", "H14", "H15", "H16", "H17"
];

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Buton Tıklama Olayları
        document.getElementById("btn-1").onclick = () => fetchData(API_URL_1, "YahooVerileri");
        document.getElementById("btn-2").onclick = () => fetchData(API_URL_2, "FinansVerileri");
        document.getElementById("btn-3").onclick = () => fetchPortfolioData(); 
        document.getElementById("btn-4").onclick = () => fetch1500GunData(); 
        document.getElementById("btn-5").onclick = () => fetchHacimData(); // 5. Buton Tıklama Olayı

        // Otomatik Güncelleme Checkbox ve Select Olayları
        document.getElementById("auto-update-cb-1").onchange = () => handleTimer(1);
        document.getElementById("interval-select-1").onchange = () => handleTimer(1);

        document.getElementById("auto-update-cb-2").onchange = () => handleTimer(2);
        document.getElementById("interval-select-2").onchange = () => handleTimer(2);

        document.getElementById("auto-update-cb-3").onchange = () => handleTimer(3);
        document.getElementById("interval-select-3").onchange = () => handleTimer(3);

        document.getElementById("auto-update-cb-4").onchange = () => handleTimer(4);
        document.getElementById("interval-select-4").onchange = () => handleTimer(4);

        document.getElementById("auto-update-cb-5").onchange = () => handleTimer(5); // 5. Checkbox Olayı
        document.getElementById("interval-select-5").onchange = () => handleTimer(5); // 5. Select Olayı
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
        status.style.color = sheetName === "YahooVerileri" ? "#217346" : "#0078d4";
    } catch (error) {
        status.innerText = "Hata: Bağlantı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

async function fetchPortfolioData() {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: Kapanış 300 Gün güncelleniyor...`;
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

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (300 Gün)`;
        status.style.color = "#5c2d91"; 
    } catch (error) {
        status.innerText = "Hata: 300 Gün bağlantısı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

async function fetch1500GunData() {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: Kapanış 1500 Gün (P Serisi) güncelleniyor...`;
        status.style.color = "#d83b01"; 

        await Excel.run(async (context) => {
            for (let item of URLS_1500) {
                let url = `${item.url}?sayfaAdi=${item.name}`;
                let response = await fetch(url);
                let data = await response.json();

                if (data.error || !data || data.length === 0) continue;

                const sheets = context.workbook.worksheets;
                let sheet = sheets.getItemOrNullObject(item.name);
                await context.sync();

                if (sheet.isNullObject) {
                    sheet = sheets.add(item.name);
                }

                sheet.getUsedRange().clear(Excel.ClearApplyTo.contents);
                const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
                range.values = data;
                range.format.autofitColumns();
            }
            await context.sync();
        });

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (1500 Gün)`;
        status.style.color = "#d83b01"; 
    } catch (error) {
        status.innerText = "Hata: 1500 Gün bağlantısı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

// 5. KART: Hacim 100 Gün Veri Çekme Fonksiyonu
async function fetchHacimData() {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: Hacim 100 Gün güncelleniyor...`;
        status.style.color = "#00828a"; 

        await Excel.run(async (context) => {
            for (let sayfaAdi of HACIM_SAYFALAR) {
                let url = `${API_URL_5}?sayfaAdi=${sayfaAdi}`;
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

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (Hacim 100 Gün)`;
        status.style.color = "#00828a"; 
    } catch (error) {
        status.innerText = "Hata: Hacim 100 Gün bağlantısı kurulamadı!";
        status.style.color = "#a4262c";
    }
}

function handleTimer(id) {
    const isChecked = document.getElementById(`auto-update-cb-${id}`).checked;
    const intervalMs = parseInt(document.getElementById(`interval-select-${id}`).value);
    const timerText = document.getElementById(`timer-text-${id}`);

    // Mevcut zamanlayıcıyı temizle
    if (id === 1) clearInterval(updateInterval1);
    else if (id === 2) clearInterval(updateInterval2);
    else if (id === 3) clearInterval(updateInterval3);
    else if (id === 4) clearInterval(updateInterval4);
    else if (id === 5) clearInterval(updateInterval5); // Eklendi

    if (isChecked) {
        let sureMetni = "";
        if (intervalMs === 86400000) {
            sureMetni = "1 Gün";
        } else if (intervalMs === 3600000) {
            sureMetni = "1 Saat";
        } else {
            sureMetni = `${intervalMs / 60000} Dakika`;
        }

        timerText.innerText = `Aktif: ${sureMetni}'da bir yenileniyor.`;
        
        // Tasarıma uygun renk ayarı
        let color = "#217346";
        if (id === 2) color = "#0078d4";
        if (id === 3) color = "#5c2d91";
        if (id === 4) color = "#d83b01";
        if (id === 5) color = "#00828a"; // 5. Kart için Teal rengi
        timerText.style.color = color; 
        
        let interval;

        if (id === 5) {
            fetchHacimData(); 
            interval = setInterval(() => { fetchHacimData(); }, intervalMs);
            updateInterval5 = interval;
        }
        else if (id === 4) {
            fetch1500GunData(); 
            interval = setInterval(() => { fetch1500GunData(); }, intervalMs);
            updateInterval4 = interval;
        } 
        else if (id === 3) {
            fetchPortfolioData(); 
            interval = setInterval(() => { fetchPortfolioData(); }, intervalMs);
            updateInterval3 = interval;
        } 
        else {
            const targetUrl = id === 1 ? API_URL_1 : API_URL_2;
            const targetSheet = id === 1 ? "YahooVerileri" : "FinansVerileri";
            
            fetchData(targetUrl, targetSheet); 
            interval = setInterval(() => { fetchData(targetUrl, targetSheet); }, intervalMs);

            if (id === 1) updateInterval1 = interval;
            else updateInterval2 = interval;
        }

    } else {
        timerText.innerText = "Zamanlayıcı Kapalı";
        timerText.style.color = "#888";
    }
}
