import { parseCatalogSnapshotCsv, parseCatalogSnapshotRecords } from '../domain/edgeAiCatalog.js';
import { readFileAutoEncoding } from './csv.js';

const getExtension = (name = '') => String(name).toLowerCase().match(/\.([^.]+)$/)?.[1] || '';

const parseBestWorkbookSheet = (sheets = []) => {
  const candidates = [];
  const errors = [];

  sheets.forEach((sheet, sheetIndex) => {
    try {
      const parsed = parseCatalogSnapshotRecords(sheet?.data || []);
      candidates.push({
        ...parsed,
        sheetName: sheet?.sheet || `Sheet${sheetIndex + 1}`,
        sheetIndex,
        score: parsed.recognizedFields.length * 100000 + parsed.items.length
      });
    } catch (error) {
      errors.push(error);
    }
  });

  if (!candidates.length) {
    throw errors[0] || new Error('介援隊コードを含むシートが見つかりません。');
  }

  return candidates.sort((left, right) => right.score - left.score || left.sheetIndex - right.sheetIndex)[0];
};

export const readCatalogSnapshotFile = async (file) => {
  if (!file) throw new Error('比較するファイルを選択してください。');
  const extension = getExtension(file.name);

  if (extension === 'csv') {
    const text = await readFileAutoEncoding(file);
    return {
      ...parseCatalogSnapshotCsv(text),
      fileName: file.name,
      fileType: 'csv',
      sheetName: ''
    };
  }

  if (extension === 'xlsx') {
    const { default: readXlsxFile } = await import('read-excel-file/browser');
    const sheets = await readXlsxFile(file);
    const parsed = parseBestWorkbookSheet(sheets);
    return {
      ...parsed,
      fileName: file.name,
      fileType: 'xlsx',
      availableSheets: sheets.map((sheet) => sheet?.sheet).filter(Boolean)
    };
  }

  if (extension === 'xls') {
    throw new Error('旧形式の.xlsには未対応です。.xlsxまたは.csvで保存してから取り込んでください。');
  }

  throw new Error('対応形式は.xlsxまたは.csvです。');
};

export { parseBestWorkbookSheet };
