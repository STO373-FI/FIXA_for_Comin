const fs = require('fs');
const xlsx = require('xlsx');

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function run() {
    console.log("엑셀 파일 파싱 시작...");
    
    // 1. 파일명 변경 적용
    const TARGET_FILE = './FIXA mediamaintenance Tool.xlsx';

    if (!fs.existsSync(TARGET_FILE)) {
        console.error(`❌ 오류: '${TARGET_FILE}' 파일을 찾을 수 없습니다.`);
        fs.writeFileSync('./data.json', JSON.stringify([{ error: "엑셀 파일 없음" }], null, 2));
        return;
    }

    const workbook = xlsx.readFile(TARGET_FILE);

    // 2. AL010 시트에서 제품번호-제품명 맵핑 딕셔너리 생성 (VLOOKUP 역할)
    const al010Sheet = workbook.Sheets['AL010'];
    const nameMap = {};
    if (al010Sheet) {
        const al010Data = xlsx.utils.sheet_to_json(al010Sheet, { header: 1 });
        al010Data.forEach(row => {
            let artNo = row[2] ? String(row[2]).trim() : ''; // C열 (인덱스 2)
            let artName = row[3] ? String(row[3]).trim() : ''; // D열 (인덱스 3)
            
            if (artNo) {
                // 8자리 이하 번호는 무조건 앞에 0을 채워 8자리로 고정 (짧은 번호 오류 방지)
                if (artNo.length <= 8) artNo = artNo.padStart(8, '0');
                nameMap[artNo] = artName;
            }
        });
        console.log("✅ AL010 시트 제품명 맵핑 딕셔너리 구성 완료.");
    } else {
        console.warn("⚠️ AL010 시트를 찾을 수 없어 제품명을 맵핑하지 못합니다.");
    }

    // 3. Query1 시트 (반드시 2번째 시트 고정) 데이터 파싱
    const query1SheetName = workbook.SheetNames[1]; // 인덱스 1 = 두 번째 시트
    if (!query1SheetName) {
        console.error("❌ 두 번째 시트(Query1)를 찾을 수 없습니다.");
        return;
    }

    const rawExcelData = xlsx.utils.sheet_to_json(workbook.Sheets[query1SheetName], { header: 1 });
    const rowsToProcess = rawExcelData.slice(1);
    
    const processedItems = [];
    const uniqueItemNos = new Set();
    const imageCache = {}; 

    rowsToProcess.forEach(row => {
        let artNo = row[3] ? String(row[3]).trim() : '';
        if (artNo && artNo.length <= 8) {
            artNo = artNo.padStart(8, '0');
        }
        
        if (artNo && artNo !== 'Empty') uniqueItemNos.add(artNo);

        processedItems.push({
            mediaType: row[1] ? String(row[1]).trim() : '',
            area: row[2] ? String(row[2]).trim() : '',
            itemNumber: artNo || 'Empty',
            itemName: nameMap[artNo] || '-', // 맵핑된 제품명 삽입 (없으면 '-')
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
