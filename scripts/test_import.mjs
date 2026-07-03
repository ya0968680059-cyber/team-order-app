import fs from 'fs/promises';
import path from 'path';

async function run() {
  const file = path.resolve(process.cwd(), 'sample-menu.csv');
  try {
    const data = await fs.readFile(file, 'utf8');
    const lines = data.split(/\r?\n/).filter(Boolean);
    if (lines.length < 2) {
      console.error('CSV 內容不足');
      process.exit(1);
    }
    const header = lines[0].split(',').map(h => h.trim());
    const rows = lines.slice(1).map(line => {
      const cols = line.split(',');
      const obj = {};
      header.forEach((h, i) => { obj[h] = (cols[i] || '').trim(); });
      return obj;
    });

    const imported = rows.map((row, index) => {
      const name = String(row['品名'] ?? row['name'] ?? row['商品名稱'] ?? '').trim();
      const price = Number(row['價格'] ?? row['price'] ?? row['價錢'] ?? 0);
      const category = String(row['分類'] ?? row['category'] ?? row['類別'] ?? '其他').trim();
      if (!name) return null;
      return { id: Date.now() + index, name, price: Number.isFinite(price) ? price : 0, category: category || '其他' };
    }).filter(Boolean);

    const storeName = path.basename(file).replace(/\.[^.]+$/, '') || '匯入菜單';

    console.log('=== 模擬匯入結果 ===');
    console.log('來源檔名（去副檔名）作為菜單名稱：', storeName);
    console.log('解析項目：');
    console.log(JSON.stringify(imported, null, 2));
  } catch (err) {
    console.error('讀取或解析失敗', err);
    process.exit(1);
  }
}

run();
