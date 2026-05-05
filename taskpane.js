async function fetchPortfolioData() {
    const status = document.getElementById("status");
    try {
        status.innerText = `İşlem: 18 Sayfa Tek Pakette Çekiliyor...`;
        status.style.color = "#5c2d91";

        // TEK BİR FETCH (18 isteğin yerine tek bir hamle)
        const response = await fetch(API_URL_3);
        const allData = await response.json();

        if (allData.error) throw new Error(allData.error);

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            // Gelen paketteki her bir sayfa için
            for (let sayfaAdi in allData) {
                let data = allData[sayfaAdi];
                
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

        status.innerText = `Son Güncelleme: ${new Date().toLocaleTimeString()} (Toplu Aktarım Başarılı)`;
        status.style.color = "#107c41";
    } catch (error) {
        console.error(error);
        status.innerText = "Hata: Toplu paket aktarılamadı!";
        status.style.color = "#a4262c";
    }
}
