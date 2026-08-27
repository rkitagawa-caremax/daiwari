export const normalizeCode = (value) => {
  if (!value) return '';
  return value
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (character) => String.fromCharCode(character.charCodeAt(0) - 0xFEE0))
    .replace(/[-\s]/g, '')
    .trim()
    .toUpperCase();
};
