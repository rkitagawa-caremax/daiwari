export const readFileAutoEncoding = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (event) => {
    const buffer = event.target.result;
    const bytes = new Uint8Array(buffer);
    const encoding = (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF)
      ? 'UTF-8'
      : 'Shift_JIS';
    try {
      const text = new TextDecoder(encoding).decode(buffer).replace(/^\uFEFF/, '');
      resolve(text);
    } catch (error) {
      reject(error);
    }
  };
  reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'));
  reader.readAsArrayBuffer(file);
});

export const parseCSVLine = (line) => {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let index = 0; index < line.length; index++) {
    const character = line[index];
    if (character === '"') {
      if (inQuotes && line[index + 1] === '"') {
        current += '"';
        index++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (character === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += character;
    }
  }
  result.push(current.trim());
  return result;
};
