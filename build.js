const fs = require('fs');
const xlsx = require('xlsx');

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function run() {
    console.log("엑셀 파일 파싱 시작...");
    
    if (!fs.existsSync('./FIXA data.xlsx')) {
        console.error("❌ 오류: 'FIXA data.xlsx' 파일을 찾을 수 없습니다.");
        fs.writeFileSync('./data.json', JSON.stringify([{ error: "엑셀 파일 없음" }], null, 2));
        return;
    }

    const workbook = xlsx.readFile('./FIXA data.xlsx');
    const sheetName = workbook.SheetNames[0];
    const rawExcelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    const rowsToProcess = rawExcelData.slice(1);
    const processedItems = [];
    const uniqueItemNos = new Set();
    const imageCache = {}; 

    rowsToProcess.forEach(row => {
        let artNo = row[3] ? String(row[3]).trim() : '';
        if (artNo && artNo.length >= 6 && artNo.length <= 8) {
            artNo = artNo.padStart(8, '0');
        }
        
        if (artNo && artNo !== 'Empty') uniqueItemNos.add(artNo);

        processedItems.push({
            mediaType: row[1] ? String(row[1]).trim() : '',
            area: row[2] ? String(row[2]).trim() : '',
            itemNumber: artNo || 'Empty',
            ssd: row[7] ? String(row[7]).trim() : '',
            eds: row[8] ? String(row[8]).trim() : ''
        });
    });

    console.log(`총 ${uniqueItemNos.size}개의 고유 제품 이미지 검색 시작...`);

    let count = 0;
    for (const itemNo of uniqueItemNos) {
        count++;
        let imgUrl = 'No Image';
        try {
            const res = await fetch(`https://sik.search.blue.cdtapps.com/kr/ko/search-box?q=${itemNo}`);
            if (res.ok) {
                const data = await res.json();
                imgUrl = findBestImage(data) || 'No Image';
            }
        } catch (e) {
            console.error(`${itemNo} 조회 실패:`, e.message);
        }
        
        imageCache[itemNo] = imgUrl;
        console.log(`[${count}/${uniqueItemNos.size}] ${itemNo} 이미지 수집 완료`);
        
        await sleep(100); 
    }

    const finalData = processedItems.map(item => ({
        ...item,
        imageUrl: imageCache[item.itemNumber] || 'No Image'
    }));

    fs.writeFileSync('./data.json', JSON.stringify(finalData, null, 2));
    console.log("🎉 data.json 파일 생성 완료!");
}

function findBestImage(obj) {
    let bestUrl = null;
    function traverse(o) {
        if (!o || typeof o !== 'object') return;
        for (let key in o) {
            if (typeof o[key] === 'string') {
                let val = o[key].toLowerCase();
                if (val.match(/\.(jpg|jpeg|png|webp|avif)(\?.*)?$/) && !val.includes('.svg')) {
                    let actualUrl = o[key];
                    if (actualUrl.startsWith('//')) actualUrl = 'https:' + actualUrl;
                    else if (actualUrl.startsWith('/')) actualUrl = 'https://www.ikea.com' + actualUrl;
                    if (!bestUrl) bestUrl = actualUrl;
                    else if (actualUrl.includes('s5') || actualUrl.includes('main')) bestUrl = actualUrl;
                }
            } else if (typeof o[key] === 'object') traverse(o[key]);
        }
    }
    traverse(obj);
    return bestUrl;
}

run();
