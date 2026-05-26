import XLSX from 'xlsx-js-style';

const GAS_URL = 'https://script.google.com/macros/s/AKfycbwpqbW9sHqUZIUwGf4naJN_ZH0aCXvGEYSdqBN6mftyFVZGAofjrnMPfutcO5maBc4/exec';

const parseNumeric = (val: any): number => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  
  let str = val.toString().trim();
  if (!/[0-9]/.test(str)) return 0;

  if (str.includes('+')) {
    const parts = str.split('+');
    return parts.reduce((acc, part) => acc + parseNumeric(part), 0);
  }

  str = str.replace(/\s/g, '');
  str = str.replace(/[^0-9.,-]/g, '');

  if (str.includes(',') && str.includes('.')) {
    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');
    if (lastComma > lastDot) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    const commaCount = (str.match(/,/g) || []).length;
    if (commaCount > 1) {
      str = str.replace(/,/g, '');
    } else {
      str = str.replace(',', '.');
    }
  } else if (str.includes('.')) {
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) {
      str = str.replace(/\./g, '');
    } else {
      const parts = str.split('.');
      if (parts[1] && parts[1].length === 3) {
        str = str.replace(/\./g, '');
      }
    }
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};

async function checkDiff() {
  try {
    const res = await fetch(`${GAS_URL}?action=download&gid=1476570479`);
    const result: any = await res.json();
    if (result.status === 'success' && result.base64) {
      const byteCharacters = Buffer.from(result.base64, 'base64');
      const workbook = XLSX.read(byteCharacters, { type: 'buffer' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rawAOA: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      rawAOA.forEach((row, idx) => {
        if (idx === 0) return;
        const id = row[0];
        if (!id) return;

        const valRaw = row[16];
        if (valRaw !== undefined && valRaw !== null && valRaw !== '') {
          const parsed = parseNumeric(valRaw);
          // Standard JS parseFloat
          const standardFloat = parseFloat(valRaw.toString().replace(/[^0-9.]/g, '')) || 0;
          
          if (parsed !== valRaw && typeof valRaw === 'string') {
            console.log(`ID: ${id}, Row Raw: "${valRaw}", Parsed: ${parsed}`);
          }
        }
      });
    }
  } catch (err) {
    console.error(err);
  }
}

checkDiff();
