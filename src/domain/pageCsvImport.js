import { normalizeCode } from './productCodes.js';
import { buildDefaultPanels } from './panels.js';
import { fillPanelArea, findFirstPlaceableIndex, getSpansFromSizeTypeRobust } from './panelLayout.js';

// 台割ページ CSV 取り込みの純粋ロジック。
// App.jsx 側は「ファイル読込 → parsePageCsvRows → buildImportedSheets → 永続化 → buildPageCsvImportReport」の順に呼ぶ。
// 進捗表示 (setProgress*) と UI への yield は onProgress コールバック経由で呼び出し側が行う。

export const PAGE_CSV_MIN_COLUMNS = 6;
export const PAGE_CSV_PROGRESS_INTERVAL = 50;
export const IMAGE_MATCH_THRESHOLD = 90;
const DEFAULT_SIZE_TYPE = '1/16（1コマ）';

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

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row.trim()) continue;

    if (i % PAGE_CSV_PROGRESS_INTERVAL === 0) {
      await onProgress?.(i, rows.length);
    }

    const cols = parseLine(row);
    if (cols.length < PAGE_CSV_MIN_COLUMNS) continue;

    // カラム定義: 0:ジャンル, 1:ページ数, 2:追番, 3:コマ番号, 4:介援隊コード, 5:コマ数, 6:ダミーラベル種別, 7:テキスト情報, 8:座標(無視), 9:コマID
    const genreLabel = cols[0];
    const pageNum = parseInt(cols[1], 10);
    const panelNumRaw = parseInt(cols[2], 10);
    const frameNumRaw = parseInt(cols[3], 10);
    const frameNum = (!isNaN(frameNumRaw) && frameNumRaw > 0)
      ? frameNumRaw
      : ((!isNaN(panelNumRaw) && panelNumRaw > 0) ? panelNumRaw : NaN);
    const panelNum = (!isNaN(panelNumRaw) && panelNumRaw > 0)
      ? panelNumRaw
      : ((!isNaN(frameNumRaw) && frameNumRaw > 0) ? frameNumRaw : NaN);
    const codeVal = cols[4] === 'ダミーコマ' ? '' : (cols[4] || '').trim();
    const isDummyMarker = cols[4] === 'ダミーコマ';
    const sizeVal = cols[5] || DEFAULT_SIZE_TYPE;
    const textVal = (cols[7] || '').trim();
    // J列: コマID（台割には反映しない。介援隊コードに紐づけてデータとして保持）
    const panelIdVal = (cols[9] || '').trim();

    const isFixed = !isNaN(frameNum) && frameNum > 0;

    // ページ番号は必須。コマ番号か追番のどちらかは必須。
    if (isNaN(pageNum)) continue;
    if (!isFixed && (isNaN(panelNum))) continue;

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
      const dummyLabel = (cols[6] || '').trim();
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

    sheetUpdates[pageIndex].contentItems.push({
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
        panelId: panelIdVal || null  // J列から読み込んだコマID（台割には非表示）
      }
    });
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

      // 前のコマ配置後の occupied を考慮し、先頭から最初に配置可能な位置を探す
      // コマ番号 (item.order) を開始座標のヒントとして使用
      // ユーザーの要件: コマ番号1→X1Y1, 2→X2Y1 ... 等。
      // findFirstPlaceableIndex に第4引数として (order-1) を渡す。
      const startCandidate = (item.order > 0 && item.order <= 16) ? item.order - 1 : 0;
      const resolvedStartIdx = findFirstPlaceableIndex(rowSpan, colSpan, occupied, startCandidate);

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
      importSummary.autoSuccess++;
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
