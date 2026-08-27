const INVALID_FILENAME_CHARACTERS = /[<>:"/\\|?*]/g;

export const sanitizePdfFilenamePart = (value, fallback = '未設定') => {
  const withoutControlCharacters = Array.from(String(value || ''))
    .filter((character) => character.charCodeAt(0) >= 32)
    .join('');
  const sanitized = withoutControlCharacters
    .replace(INVALID_FILENAME_CHARACTERS, '_')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim();
  return sanitized || fallback;
};

export const buildPdfExportPlan = ({ sheets = [], selectedSheetIds, genres = [] } = {}) => {
  const selectedIds = selectedSheetIds instanceof Set
    ? selectedSheetIds
    : new Set(Array.isArray(selectedSheetIds) ? selectedSheetIds : []);
  const genreLabelById = new Map(genres.map((genre) => [genre.id, genre.label]));

  const pages = (Array.isArray(sheets) ? sheets : [])
    .map((sheet, index) => ({
      sheet,
      pageNumber: index + 1,
      genreLabel: genreLabelById.get(sheet?.genre) || '未設定'
    }))
    .filter(({ sheet }) => sheet?.id && selectedIds.has(sheet.id));

  if (pages.length === 0) {
    return { pages: [], filename: '' };
  }

  if (pages.length === 1) {
    const page = pages[0];
    const genre = sanitizePdfFilenamePart(page.genreLabel);
    return { pages, filename: `Page${page.pageNumber}_${genre}.pdf` };
  }

  const uniqueGenres = [];
  pages.forEach(({ genreLabel }) => {
    const sanitizedGenre = sanitizePdfFilenamePart(genreLabel);
    if (!uniqueGenres.includes(sanitizedGenre)) uniqueGenres.push(sanitizedGenre);
  });
  const genreFilename = uniqueGenres.join('・').slice(0, 120) || '複数ジャンル';
  return { pages, filename: `${genreFilename}.pdf` };
};
