export const normalizeCode = (value) => {
  if (!value) return '';
  return value
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (character) => String.fromCharCode(character.charCodeAt(0) - 0xFEE0))
    .replace(/[-\s]/g, '')
    .trim()
    .toUpperCase();
};

// コマに記載された介援隊コードをすべて取り出す (重複は除き、記載順を保つ)。
// 「E1760 / E1761」「261-E0340」のような表記を想定し、
// 商品コードの形 (英1〜2字+数字3〜5桁) に一致するトークンだけを拾う。
export const extractProductCodes = (...values) => {
  const codes = [];
  const seen = new Set();
  values.forEach((value) => {
    if (!value) return;
    const normalized = String(value)
      .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (character) => String.fromCharCode(character.charCodeAt(0) - 0xFEE0))
      .toUpperCase();
    // カタログ番号の前置き (261- など) がコードの数字とつながらないよう、区切りは残したまま走査する
    const matches = normalized.match(/[A-Z]{1,2}\d{3,5}/g) || [];
    matches.forEach((code) => {
      if (seen.has(code)) return;
      seen.add(code);
      codes.push(code);
    });
  });
  return codes;
};
