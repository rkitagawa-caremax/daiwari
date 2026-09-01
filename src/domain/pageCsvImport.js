import { normalizeCode } from './productCodes.js';
import { buildDefaultPanels } from './panels.js';
import {
  canPlacePanelAt,
  fillPanelArea,
  findFirstPlaceableIndex,
  getSpansFromSizeTypeRobust
} from './panelLayout.js';

// 台割ページ CSV 取り込みの純粋ロジック。
// App.jsx 側は「ファイル読込 → parsePageCsvRows → buildImportedSheets → 永続化 → buildPageCsvImportReport」の順に呼ぶ。
// 進捗表示 (setProgress*) と UI への yield は onProgress コールバック経由で呼び出し側が行う。

export const PAGE_CSV_MIN_COLUMNS = 6;
export const PAGE_CSV_PROGRESS_INTERVAL = 50;
export const IMAGE_MATCH_THRESHOLD = 90;
const DEFAULT_SIZE_TYPE = '1/16（1コマ）';

const normalizeHeader = (value = '') => String(value)
  .normalize('NFKC')
  .trim()
  .toLowerCase()
  .replace(/[\s_-]+/g, '');

const PAGE_CSV_HEADER_ALIASES = Object.freeze({
  genre: ['ジャンル', '掲載ブロック'],
  page: ['ページ数', 'ページ番号', 'ページ', 'page', 'pageno', '頁'],
  order: ['追番', '順番', 'order'],
  frame: ['コマ番号', '枠番号', 'frameno'],
  code: ['介援隊コード', '介援隊cd', '商品コード', 'code'],
  size: ['コマ数', 'コマサイズ', 'サイズ', 'sizetype'],
  dummyLabel: ['ダミーラベル種別', 'ダミー種別'],
  kind: ['コマ種別', '種別', 'kind'],
  text: ['テキスト情報', 'テキスト', 'text'],
  coordinate: ['座標', 'coordinate', 'position'],
  panelId: ['コマid', 'panelid'],
  xPos: ['xpos', 'x座標'],
  yPos: ['ypos', 'y座標'],
  catalogName: ['掲載名', '商品名']
});

const LEGACY_PAGE_CSV_COLUMNS = Object.freeze({
  genre: 0,
  page: 1,
  order: 2,
  frame: 3,
  code: 4,
  size: 5,
  dummyLabel: 6,
  kind: -1,
  text: 7,
  coordinate: 8,
  panelId: 9,
  xPos: 10,
  yPos: 11,
  catalogName: -1
});

const findHeaderIndex = (headers, aliases) => {
  const normalizedAliases = new Set(aliases.map(normalizeHeader));
  return headers.findIndex((header) => normalizedAliases.has(normalizeHeader(header)));
};

export const resolvePageCsvColumns = (headers = []) => {
  const matched = Object.fromEntries(Object.entries(PAGE_CSV_HEADER_ALIASES).map(([key, aliases]) => (
    [key, findHeaderIndex(headers, aliases)]
  )));
  const legacyLayout = matched.genre === 0 && matched.page === 1 && matched.code === 4 && matched.size === 5;
  return Object.fromEntries(Object.keys(PAGE_CSV_HEADER_ALIASES).map((key) => [
    key,
    matched[key] >= 0 ? matched[key] : (legacyLayout ? LEGACY_PAGE_CSV_COLUMNS[key] : -1)
  ]));
};

const valueAt = (cols, index) => index >= 0 ? String(cols[index] ?? '') : '';

const parseGridPosition = (value) => {
  const parsed = Number.parseInt(String(value || '').normalize('NFKC'), 10);
  return Number.isInteger(parsed) && parsed >= 1 && parsed <= 4 ? parsed : null;
};

const parseGridCoordinate = (value = '') => {
  const match = String(value).normalize('NFKC').match(/X\s*([1-4])\s*Y\s*([1-4])/i);
  return match ? { xPos: Number(match[1]), yPos: Number(match[2]) } : { xPos: null, yPos: null };
};

// 全角英数字を半角に変換＆小文字化
export const normalizeImageSearchText = (value) => (
  value.replace(/[Ａ-Ｚａ-ｚ０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).toLowerCase()
);

// 画像マッチングに必要な情報だけを持つ軽量リストにする (画像データ本体は持たない)。
export const buildSearchableImages = (images = []) => images.map((img) => ({
  id: img.id,
  name: img.name || ''
}));

// 介援隊コードに最も近い画像をスコアリングで探す。
// 1. 完全一致 (100, 拡張子なし) / 2. 数値トークン完全一致 (80 or 60) / 3. 記号除去一致 (50) / 4. 包含一致 (30, 非数値のみ)
export const findBestImageMatch = (targetCode, searchableImages = []) => {
  const searchCode = normalizeImageSearchText(targetCode);
  const isNumericSearch = /^\d+$/.test(searchCode);
  const searchNum = isNumericSearch ? parseInt(searchCode, 10) : null;

  let bestMatchImg = null;
  let bestScore = 0;

  for (const img of searchableImages) {
    const normImgName = normalizeImageSearchText(img.name);
    const stem = normImgName.lastIndexOf('.') !== -1
      ? normImgName.substring(0, normImgName.lastIndexOf('.'))
      : normImgName;

    let score = 0;

    if (stem === searchCode) {
      score = 100;
    } else if (isNumericSearch) {
      const numTokens = stem.match(/\d+/g);
      if (numTokens) {
        if (numTokens.some((t) => parseInt(t, 10) === searchNum)) {
          if (numTokens.length === 1 && stem.replace(/\d+/g, '').length < stem.length) {
            score = 80;
          } else {
            score = 60;
          }
        }
      }
    } else {
      const clean = (s) => s.replace(/[^a-z0-9]/g, '');
      if (clean(stem) === clean(searchCode)) {
        score = 50;
      } else if (stem.includes(searchCode)) {
        score = 30;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatchImg = img;
      if (score === 100) break; // 完全一致なら即決
    }
  }

  return { bestMatchImg, bestScore };
};

// CSV 行 (1行目はヘッダー) をページごとの更新情報にまとめる。
// 戻り値: { sheetUpdates: { [pageIndex]: { genre, contentItems, matchDetails?, unmatchedCodes? } }, maxPageIndex }
// onProgress(rowIndex, totalRows) は PAGE_CSV_PROGRESS_INTERVAL 行ごとに await される (UI 更新用)。
export const parsePageCsvRows = async (rows, { parseLine, images = [], genres = [], onProgress } = {}) => {
  const searchableImages = buildSearchableImages(images);
  const sheetUpdates = {};
  let maxPageIndex = -1;
  const headers = rows.length > 0 ? parseLine(rows[0]) : [];
  const indexes = resolvePageCsvColumns(headers);

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row.trim()) continue;

    if (i % PAGE_CSV_PROGRESS_INTERVAL === 0) {
      await onProgress?.(i, rows.length);
    }

    const cols = parseLine(row);
    if (cols.length < PAGE_CSV_MIN_COLUMNS) continue;

    const genreLabel = valueAt(cols, indexes.genre).trim();
    const pageNum = parseInt(valueAt(cols, indexes.page), 10);
    const panelNumRaw = parseInt(valueAt(cols, indexes.order), 10);
    const frameNumRaw = parseInt(valueAt(cols, indexes.frame), 10);
    const frameNum = (!isNaN(frameNumRaw) && frameNumRaw > 0)
      ? frameNumRaw
      : ((!isNaN(panelNumRaw) && panelNumRaw > 0) ? panelNumRaw : NaN);
    const panelNum = (!isNaN(panelNumRaw) && panelNumRaw > 0)
      ? panelNumRaw
      : ((!isNaN(frameNumRaw) && frameNumRaw > 0) ? frameNumRaw : NaN);
    const rawCode = valueAt(cols, indexes.code).trim();
    const kindVal = valueAt(cols, indexes.kind).trim();
    const codeVal = rawCode === 'ダミーコマ' ? '' : rawCode;
    const isNonProductKind = !!kindVal && !/^(商品|製品)$/i.test(kindVal);
    const isDummyMarker = rawCode === 'ダミーコマ' || (!codeVal && isNonProductKind);
    const sizeVal = valueAt(cols, indexes.size).trim() || DEFAULT_SIZE_TYPE;
    const textVal = valueAt(cols, indexes.text).trim();
    const panelIdVal = valueAt(cols, indexes.panelId).trim();
    const coordinate = parseGridCoordinate(valueAt(cols, indexes.coordinate));
    const xPos = parseGridPosition(valueAt(cols, indexes.xPos)) || coordinate.xPos;
    const yPos = parseGridPosition(valueAt(cols, indexes.yPos)) || coordinate.yPos;
    const positionIndex = xPos && yPos ? (yPos - 1) * 4 + xPos - 1 : null;

    const isFixed = !isNaN(frameNum) && frameNum > 0;

    // ページ番号は必須。コマ番号・追番・明示座標のいずれかは必須。
    if (isNaN(pageNum)) continue;
    if (!isFixed && isNaN(panelNum) && positionIndex === null) continue;

    const pageIndex = pageNum - 1;
    if (pageIndex > maxPageIndex) maxPageIndex = pageIndex;

    if (!sheetUpdates[pageIndex]) {
      sheetUpdates[pageIndex] = { genre: null, contentItems: [] };
    }

    const genreObj = genres.find((g) => g.label === genreLabel);
    if (genreObj) sheetUpdates[pageIndex].genre = genreObj.id;

    let targetImageId = null;
    let targetCode = null;
    let targetLabel = null;
    let isText = false;

    // 優先順位: 1.テキスト 2.ダミーコマ 3.コード(画像)
    if (textVal) {
      targetLabel = 'テキスト';
      isText = true;
      targetCode = codeVal ? normalizeCode(codeVal) : null;
    } else if (isDummyMarker) {
      // ダミーコマの種別を判定: cols[6]（ラベル列）、cols[5]（コマ数列）、cols[4]の順にチェック
      const dummyLabel = valueAt(cols, indexes.dummyLabel).trim() || kindVal;
      if (dummyLabel === 'タイトル' || sizeVal.includes('タイトル')) {
        targetLabel = 'タイトル';
      } else if (dummyLabel === '埋草' || sizeVal.includes('埋草')) {
        targetLabel = '埋草';
      } else if (dummyLabel === '新規商品' || dummyLabel === '新規商品未確定') {
        targetLabel = '新規商品未確定';
      } else if (dummyLabel) {
        // CSVに記載のあるラベルをそのまま使用
        targetLabel = dummyLabel;
      } else {
        targetLabel = '新規商品未確定';
      }
    } else if (codeVal) {
      const normalizedToken = codeVal.normalize('NFKC').replace(/[-\s]/g, '').toUpperCase();
      const isLikelyProductCode = /^[A-Z]{1,2}\d{3,5}$/.test(normalizedToken);

      if (!isLikelyProductCode) {
        targetLabel = codeVal;
      } else {
        targetCode = normalizedToken;

        const { bestMatchImg, bestScore } = findBestImageMatch(targetCode, searchableImages);

        // スコア閾値以上のマッチングのみを有効とする
        if (bestMatchImg && bestScore >= IMAGE_MATCH_THRESHOLD) {
          targetImageId = bestMatchImg.id || null;
          // マッチング詳細を記録
          sheetUpdates[pageIndex].matchDetails = sheetUpdates[pageIndex].matchDetails || [];
          sheetUpdates[pageIndex].matchDetails.push({
            code: targetCode,
            imageName: bestMatchImg.name,
            score: bestScore,
            csvRow: i + 1
          });
        } else if (targetCode) {
          // マッチしなかったコードを記録
          sheetUpdates[pageIndex].unmatchedCodes = sheetUpdates[pageIndex].unmatchedCodes || [];
          sheetUpdates[pageIndex].unmatchedCodes.push({
            code: targetCode,
            csvRow: i + 1,
            bestScore: bestScore,
            bestMatch: bestMatchImg ? bestMatchImg.name : 'なし'
          });
        }
      }
    }

    const contentItem = {
      isFixed: isFixed,
      frameNo: isFixed ? frameNum : -1,
      order: isNaN(panelNum) ? 9999 : panelNum,
      data: {
        code: targetCode,
        image: null,
        imageId: targetImageId,
        label: targetLabel,
        sizeType: sizeVal,
        text: textVal,
        isText: isText,
        panelId: panelIdVal || null
      }
    };
    if (positionIndex !== null) contentItem.positionIndex = positionIndex;
    sheetUpdates[pageIndex].contentItems.push(contentItem);
  }

  return { sheetUpdates, maxPageIndex };
};

export const createPageCsvImportSummary = () => ({
  total: 0,
  fixedSuccess: 0,
  fixedFailed: 0,
  autoSuccess: 0,
  autoFailed: 0,
  details: [],
  matchedImages: [], // マッチした画像の詳細
  notMatchedCodes: [], // マッチしなかったコードのリスト
  imageUsageCount: {} // 画像の使用回数
});

// 既存シートに sheetUpdates を適用した新しいシート配列とサマリーを返す (元の配列は変更しない)。
// 足りないページは generateId() で採番して末尾に補う。
// onProgress(pageIndex, finalPageCount) は更新対象ページごとに await される。
export const buildImportedSheets = async ({
  sheets = [],
  sheetUpdates = {},
  finalPageCount = 0,
  generateId,
  now = () => Date.now(),
  onProgress
} = {}) => {
  const localSheets = [...sheets];

  // 足りないページをパディング
  while (localSheets.length < finalPageCount) {
    localSheets.push({
      id: generateId(),
      createdAt: { seconds: now() / 1000 },
      genre: 'none',
      panels: buildDefaultPanels()
    });
  }

  const importSummary = createPageCsvImportSummary();

  for (let i = 0; i < finalPageCount; i++) {
    const update = sheetUpdates[i];
    if (!update) continue;

    await onProgress?.(i, finalPageCount);

    const currentSheet = { ...localSheets[i] };
    if (update.genre) currentSheet.genre = update.genre;

    const newPanels = buildDefaultPanels();

    const occupied = new Set();

    // === コマ番号順 → 先頭空きスロットへ順次配置 ===
    // コマ番号順にソートし、各コマを「前のコマ配置後の最初の空き位置」に配置する。
    // 例: コマ1(1/8横 2コマ) → idx=0(X1Y1),1(X2Y1)占有
    //     コマ2(1/8横 2コマ) → 次の空き=idx=2(X3Y1),3(X4Y1)占有
    //     コマ3              → 次の空き=idx=4(X1Y2)から
    const allItems = [...update.contentItems].sort((a, b) => {
      const aKey = a.frameNo > 0 ? a.frameNo : (a.order > 0 ? a.order : Number.MAX_SAFE_INTEGER);
      const bKey = b.frameNo > 0 ? b.frameNo : (b.order > 0 ? b.order : Number.MAX_SAFE_INTEGER);
      return aKey - bKey;
    });

    for (const item of allItems) {
      importSummary.total++;
      const { data } = item;
      const { r: rowSpan, c: colSpan } = getSpansFromSizeTypeRobust(data.sizeType);

      // X_POS / Y_POS または「X1Y1」座標があれば、その位置を最優先する。
      // 重複・グリッド外の場合だけ、従来どおりコマ番号順の空き位置へ退避する。
      const hasExplicitPosition = Number.isInteger(item.positionIndex);
      let resolvedStartIdx = -1;
      if (hasExplicitPosition) {
        if (canPlacePanelAt(item.positionIndex, rowSpan, colSpan, occupied)) {
          resolvedStartIdx = item.positionIndex;
          importSummary.fixedSuccess++;
        } else {
          importSummary.fixedFailed++;
        }
      }
      if (resolvedStartIdx === -1) {
        const startCandidate = (item.order > 0 && item.order <= 16) ? item.order - 1 : 0;
        resolvedStartIdx = findFirstPlaceableIndex(rowSpan, colSpan, occupied, startCandidate);
      }

      if (resolvedStartIdx === -1) {
        importSummary.autoFailed++;
        importSummary.details.push(`・ページ ${i + 1}: 「${data.code || data.label || data.text || '不明'}」（${data.sizeType}）を配置できませんでした。`);
        continue;
      }

      newPanels[resolvedStartIdx] = {
        ...newPanels[resolvedStartIdx],
        ...data,
        rowSpan,
        colSpan,
        hidden: false
      };
      fillPanelArea(newPanels, resolvedStartIdx, rowSpan, colSpan, occupied);
      // 明示位置が衝突して空きへ退避した場合も、自動配置として記録する。
      if (!hasExplicitPosition || resolvedStartIdx !== item.positionIndex) importSummary.autoSuccess++;
    }

    currentSheet.panels = newPanels;
    localSheets[i] = currentSheet;

    // マッチング詳細をサマリーに集約
    if (update.matchDetails) {
      update.matchDetails.forEach((detail) => {
        importSummary.matchedImages.push({
          page: i + 1,
          ...detail
        });
        // 画像使用回数をカウント
        const imgKey = detail.imageName;
        importSummary.imageUsageCount[imgKey] = (importSummary.imageUsageCount[imgKey] || 0) + 1;
      });
    }

    // 未マッチコードをサマリーに集約
    if (update.unmatchedCodes) {
      update.unmatchedCodes.forEach((unmatched) => {
        importSummary.notMatchedCodes.push({
          page: i + 1,
          ...unmatched
        });
      });
    }
  }

  return { localSheets, importSummary };
};

// 取り込み完了時にユーザーへ表示するレポート文字列を組み立てる。
export const buildPageCsvImportReport = (importSummary) => {
  // 重複使用されている画像を抽出
  const duplicateImages = Object.entries(importSummary.imageUsageCount)
    .filter(([, count]) => count > 1)
    .map(([name, count]) => `${name} (${count}回)`);

  return [
    `取り込みが完了しました。(全${importSummary.total}件)`,
    `・指定通りの配置: ${importSummary.fixedSuccess}件`,
    `・空きへの自動配置: ${importSummary.autoSuccess}件`,
    importSummary.fixedFailed > 0 ? `・指定位置が重複または不足により移動: ${importSummary.fixedFailed}件` : '',
    importSummary.autoFailed > 0 ? `・スペース不足で配置失敗: ${importSummary.autoFailed}件` : '',
    '',
    `【画像マッチング結果】`,
    `・マッチング成功: ${importSummary.matchedImages.length}件`,
    `・マッチング失敗: ${importSummary.notMatchedCodes.length}件`,
    importSummary.notMatchedCodes.length > 0 ? `\n【マッチしなかったコード】` : '',
    ...importSummary.notMatchedCodes.slice(0, 15).map((item) =>
      `・P${item.page} 行${item.csvRow}: ${item.code} (ベストマッチ: ${item.bestMatch}, スコア: ${item.bestScore})`
    ),
    importSummary.notMatchedCodes.length > 15 ? `...他 ${importSummary.notMatchedCodes.length - 15} 件` : '',
    duplicateImages.length > 0 ? `\n【重複使用されている画像】` : '',
    ...duplicateImages.slice(0, 10).map((item) => `・${item}`),
    duplicateImages.length > 10 ? `...他 ${duplicateImages.length - 10} 件` : '',
    importSummary.details.length > 0 ? "\n【未配置の項目】\n" + importSummary.details.slice(0, 10).join('\n') + (importSummary.details.length > 10 ? '\n...他' : '') : ''
  ].filter(Boolean).join('\n');
};
