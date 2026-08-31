// ブラウザ上でテキストをファイルとしてダウンロードさせる。
// (a 要素を一時的に生成して click する従来方式をそのまま関数化)
export const downloadTextFile = (content, filename, mimeType = 'text/csv;charset=utf-8;') => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
