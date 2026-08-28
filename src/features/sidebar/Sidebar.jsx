import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  Ban,
  Check,
  CheckSquare,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  FileSpreadsheet,
  Image as ImageIcon,
  Info,
  Search,
  Trash2,
  Type,
  X
} from 'lucide-react';

import { FREE_LABEL_COLORS, GENRES } from '../../constants/layout';
import ImagePreviewModal from '../../components/dialogs/ImagePreviewModal';
import { buildSidebarImageResults } from '../../domain/sidebarImageSearch';
import { getPanelFreeLabels } from '../../domain/panels';
import {
  clearActiveNativeDragPayload,
  getDragPayload,
  isDropEventHandled,
  markDropEventHandled,
  setDragPayload
} from '../../lib/dragPayload';

const FreeLabelPreview = ({ item }) => {
  const labels = getPanelFreeLabels(item);
  if (labels.length === 0) return null;

  return labels.map((label, index) => {
    const rawColorIndex = Number.parseInt(label.colorIndex, 10);
    const colorIndex = Number.isNaN(rawColorIndex) ? index : rawColorIndex;
    const color = FREE_LABEL_COLORS[((colorIndex % FREE_LABEL_COLORS.length) + FREE_LABEL_COLORS.length) % FREE_LABEL_COLORS.length];
    const rawX = Number(label.x);
    const rawY = Number(label.y);
    const x = Number.isFinite(rawX) ? Math.min(100, Math.max(0, rawX)) : 50;
    const y = Number.isFinite(rawY) ? Math.min(100, Math.max(0, rawY)) : 50;

    return (
      <span
        key={label.id || `free-label-${index}`}
        className="pointer-events-none absolute z-10 max-w-[90%] -translate-x-1/2 -translate-y-1/2 truncate rounded px-1 py-0.5 text-[8px] font-bold leading-tight text-white shadow"
        style={{
          left: `${x}%`,
          top: `${y}%`,
          backgroundColor: color.bg,
          border: `1px solid ${color.border}`
        }}
      >
        {label.text || 'ラベル'}
      </span>
    );
  });
};

const DUMMY_OPTIONS = [
  { color: 'var(--dummy-red)', text: 'var(--dummy-text-color)', label: 'タイトル' },
  { color: 'var(--dummy-green)', text: 'var(--dummy-text-color)', label: '埋草' },
  { color: 'var(--dummy-volt)', text: 'var(--dummy-text-color)', label: 'テキスト', isText: true },
];

const getSheetEmptyCount = (sheet) => {
  const panels = Array.isArray(sheet?.panels) ? sheet.panels : [];
  const pureEmpty = panels.filter((panel) => !panel.hidden && !panel.image && !panel.imageId && !panel.label).length;
  const dummy = panels.filter((panel) => !panel.hidden && panel.label && panel.label !== 'タイトル' && panel.label !== '埋草' && panel.label !== 'テキスト').length;
  return pureEmpty + dummy;
};

// ... Sidebar Component ...
const Sidebar = React.memo(({
  isLocked,
  isOpen,
  isTopBarsVisible,
  width,
  setWidth,
  toggleOpen,
  images,
  onUpload,
  onDeleteImage,
  onBulkDeleteImages,
  onSearch,
  searchQuery,
  sheets,
  tempItems,
  onDeleteFromTemp,
  excludedItems,
  onDeleteFromExcluded,
  onExportExcludedCSV,
  onBulkDeleteExcluded,
  onApplyDragPayloadToTemp,
  onApplyDragPayloadToExcluded,
  onApplyDragPayloadToStock,
  onOpenAssignedImage,
  onStartPointerDrag,
  imageDataById,
  onShowQuickHelp,
  onHideQuickHelp
}) => {
  const [activeTab, setActiveTab] = useState('stock');
  const [resizing, setResizing] = useState(false);
  const [statusGenreFilter, setStatusGenreFilter] = useState('all');
  const [isImageSelectionMode, setIsImageSelectionMode] = useState(false);
  const [selectedImageIds, setSelectedImageIds] = useState(new Set());
  const [previewImage, setPreviewImage] = useState(null);
  const [excludedSearchQuery, setExcludedSearchQuery] = useState('');

  const sidebarTabHelp = {
    stock: { title: '画像', description: '画像ライブラリを表示します。画像の検索・アップロード・配置ができます。' },
    dummy: { title: 'ダミー', description: 'ダミー素材をドラッグしてコマに配置します。テキストダミーもここから使えます。' },
    status: { title: '空き', description: 'ページごとの空きコマ状況を確認できます。ジャンル絞り込みにも対応しています。' },
    excluded: { title: '除外', description: '掲載除外リストです。除外した素材の確認・復帰・削除ができます。' }
  };

  const getImageSelectionKey = useCallback((img) => {
    if (img?.searchResultKey) return img.searchResultKey;
    if (img?.id) return `id:${img.id}`;
    const data = img?.data || '';
    const head = data.slice(0, 32);
    const tail = data.slice(-32);
    return `legacy:${img?.name || ''}:${data.length}:${head}:${tail}`;
  }, []);

  useEffect(() => {
    setSelectedImageIds(new Set());
    setIsImageSelectionMode(false);
  }, [activeTab]);

  const toggleImageSelection = (selectionKey) => {
    setSelectedImageIds(prev => {
      const newSet = new Set(prev);
      if (newSet.has(selectionKey)) newSet.delete(selectionKey);
      else newSet.add(selectionKey);
      return newSet;
    });
  };

  const handleBulkDelete = () => {
    if (isLocked) return;
    const selectedTargets = filteredImages
      .filter((img) => selectedImageIds.has(getImageSelectionKey(img)))
      .map((img) => ({ id: img.id || null, data: img.data || null }));
    onBulkDeleteImages(selectedTargets);
    setIsImageSelectionMode(false);
    setSelectedImageIds(new Set());
  };

  const handleTogglePreview = useCallback((src, name = '') => {
    if (!src) return;
    setPreviewImage((prev) => {
      if (prev?.src === src) return null;
      return { src, name };
    });
  }, []);

  useEffect(() => {
    const handleMouseMove = (e) => {
      if (resizing) {
        const newWidth = Math.max(200, Math.min(600, e.clientX));
        setWidth(newWidth);
      }
    };
    const handleMouseUp = () => setResizing(false);

    if (resizing) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizing, setWidth]);

  const createDummyDragPayload = useCallback((dummy) => {
    const canvas = document.createElement('canvas');
    canvas.width = 100;
    canvas.height = 100;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#eef2ff';
    ctx.fillRect(0, 0, 100, 100);
    return {
      src: canvas.toDataURL(),
      label: dummy.label,
      isText: dummy.isText ? 'true' : ''
    };
  }, []);

  const handleDropToTemp = (e) => {
    e.preventDefault();
    const nativeDropEvent = e.nativeEvent;
    if (isDropEventHandled(nativeDropEvent)) return;
    const handled = onApplyDragPayloadToTemp?.(getDragPayload(e.dataTransfer) || {});
    if (handled) {
      markDropEventHandled(nativeDropEvent);
      clearActiveNativeDragPayload();
    }
  };

  const handleDropToStock = (e) => {
    e.preventDefault();
    const nativeDropEvent = e.nativeEvent;
    if (isDropEventHandled(nativeDropEvent)) return;
    const handled = onApplyDragPayloadToStock?.(getDragPayload(e.dataTransfer) || {});
    if (handled) {
      markDropEventHandled(nativeDropEvent);
      clearActiveNativeDragPayload();
    }
  };

  const handleDropToExcluded = (e) => {
    e.preventDefault();
    const nativeDropEvent = e.nativeEvent;
    if (isDropEventHandled(nativeDropEvent)) return;
    const handled = onApplyDragPayloadToExcluded?.(getDragPayload(e.dataTransfer) || {});
    if (handled) {
      markDropEventHandled(nativeDropEvent);
      clearActiveNativeDragPayload();
    }
  };

  // 通常時は未配置画像だけを表示し、検索中はコードが一致する配置済み画像も合成する。
  const filteredImages = useMemo(() => {
    return buildSidebarImageResults({ images, sheets, excludedItems, searchQuery });
  }, [images, sheets, excludedItems, searchQuery]);

  const activeTempItems = useMemo(() => tempItems || [], [tempItems]);

  // 除外リスト検索: code / label / originalName / text を case-insensitive で部分一致
  // 注: 一括削除 / CSV 出力ボタンは全件 (excludedItems) を対象にする (検索は表示フィルタのみ)
  const filteredExcludedItems = useMemo(() => {
    const q = excludedSearchQuery.trim().toLowerCase();
    if (!q) return excludedItems;
    return (excludedItems || []).filter((item) => {
      if (!item) return false;
      const code = (item.code || '').toLowerCase();
      const label = (item.label || '').toLowerCase();
      const name = (item.originalName || '').toLowerCase();
      const text = (typeof item.text === 'string' ? item.text : '').toLowerCase();
      return code.includes(q) || label.includes(q) || name.includes(q) || text.includes(q);
    });
  }, [excludedItems, excludedSearchQuery]);

  const statusFilteredSheets = useMemo(() => {
    return (sheets || [])
      .map((sheet, originalIndex) => ({ sheet, originalIndex }))
      .filter(({ sheet }) => statusGenreFilter === 'all' || sheet.genre === statusGenreFilter);
  }, [sheets, statusGenreFilter]);

  const statusFilteredEmptyCount = useMemo(() => {
    return statusFilteredSheets.reduce((total, { sheet }) => total + getSheetEmptyCount(sheet), 0);
  }, [statusFilteredSheets]);

  if (!isOpen) {
    return (
      <div className={`fixed left-0 ${isTopBarsVisible ? 'top-20' : 'top-0'} bottom-0 w-16 m3-surface border-r flex flex-col items-center py-6 z-20 transition-all duration-300`} style={{ borderColor: 'var(--m3-outline-variant)' }}>
        <button
          onClick={toggleOpen}
          className="m3-icon-btn-tonal p-3"
          style={{ background: 'var(--m3-primary-container)', color: 'var(--m3-on-primary-container)' }}
        >
          <ChevronRight size={24} />
        </button>
      </div>
    );
  }

  return (
    <div
      className={`fixed left-0 ${isTopBarsVisible ? 'top-20' : 'top-0'} bottom-0 m3-surface border-r flex flex-col z-20 shadow-xl transition-all duration-300 ease-in-out`}
      style={{ width, borderColor: 'var(--m3-outline-variant)' }}
    >
      <div className="flex items-start gap-2 px-3 py-2.5 border-b flex-shrink-0" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
        <div
          className="grid min-w-0 flex-1 grid-cols-2 gap-px overflow-hidden rounded-lg border"
          style={{ background: 'var(--m3-outline-variant)', borderColor: 'var(--m3-outline-variant)' }}
          role="tablist"
          aria-label="サイドパネル表示"
        >
          {[
            { id: 'stock', icon: ImageIcon, label: '画像' },
            { id: 'dummy', icon: Type, label: 'ダミー' },
            { id: 'status', icon: Info, label: '空き' },
            { id: 'excluded', icon: Ban, label: '除外' }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              data-work-action="navigation"
              onMouseEnter={(e) => {
                const help = sidebarTabHelp[tab.id];
                if (help) onShowQuickHelp?.(e, help.title, help.description);
              }}
              onMouseLeave={() => onHideQuickHelp?.()}
              role="tab"
              aria-selected={activeTab === tab.id}
              className={`flex min-h-10 min-w-0 flex-col items-center justify-center gap-0 px-2 py-1.5 transition-all duration-200 ${activeTab === tab.id ? 'relative z-10 shadow-sm' : 'hover:brightness-95'}`}
              style={activeTab === tab.id
                ? { color: 'var(--m3-on-primary-container)', background: 'var(--m3-primary-container)' }
                : { color: 'var(--m3-on-surface-variant)', background: 'var(--m3-surface)' }}
            >
              <tab.icon size={16} strokeWidth={activeTab === tab.id ? 2.6 : 2} />
              <span className="truncate text-[10px] font-bold leading-tight">{tab.label}</span>
            </button>
          ))}
        </div>
        <button onClick={toggleOpen} className="m3-icon-btn mt-0.5 flex-shrink-0" title="サイドパネルを閉じる" aria-label="サイドパネルを閉じる">
          <ChevronLeft size={24} />
        </button>
      </div>

      {activeTab === 'stock' && (
        <div className="flex-shrink-0 border-b p-3" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}>
          <div className="relative group">
            <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 transition-colors" style={{ color: 'var(--m3-outline)' }} />
            <input
              type="text"
              data-work-action="navigation"
              placeholder="画像検索..."
              value={searchQuery}
              onChange={(e) => onSearch(e.target.value)}
              onMouseEnter={(e) => onShowQuickHelp?.(e, '画像検索', '画像名・介援隊コードで検索します。配置済み画像はページ番号付きで表示され、クリックすると該当ページへ移動します。')}
              onMouseLeave={() => onHideQuickHelp?.()}
              className="w-full rounded-lg py-2.5 pl-10 pr-8 text-sm transition-all"
              style={{ background: 'var(--m3-surface-container-high)', color: 'var(--m3-on-surface)' }}
            />
            {searchQuery && (
              <button
                onClick={() => onSearch('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 rounded-full p-1 transition-colors hover:bg-black/10"
                style={{ color: 'var(--m3-on-surface-variant)' }}
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>
      )}

      <div
        className="flex-1 overflow-y-auto p-3 relative"
        style={{ background: 'var(--m3-surface-container-low)' }}
        data-daiwari-dropzone-id={activeTab === 'stock' ? 'stock' : undefined}
        onDragOverCapture={(e) => e.preventDefault()}
        onDropCapture={activeTab === 'stock' ? handleDropToStock : undefined}
        onDragOver={activeTab === 'stock' ? (e) => e.preventDefault() : undefined}
        onDrop={activeTab === 'stock' ? handleDropToStock : undefined}
      >

        {activeTab === 'stock' && (
          <div className="space-y-4">
            <div className="space-y-2">
              <label
                className="flex h-24 w-full cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed transition-all group"
                style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}
                onMouseEnter={(e) => onShowQuickHelp?.(e, 'クリックしてアップロード', '画像ファイルを選択してライブラリに追加します。複数選択アップロードにも対応します。')}
                onMouseLeave={() => onHideQuickHelp?.()}
              >
                <div className="flex flex-col items-center py-2">
                  <div className="mb-2 rounded-full p-2 transition-colors" style={{ background: 'var(--m3-primary-container)' }}>
                    <ImageIcon className="h-5 w-5" style={{ color: 'var(--m3-on-primary-container)' }} />
                  </div>
                  <p className="text-xs font-bold" style={{ color: 'var(--m3-on-surface-variant)' }}>画像を追加</p>
                  <p className="mt-0.5 text-[10px]" style={{ color: 'var(--m3-outline)' }}>クリックして選択</p>
                </div>
                <input type="file" className="hidden" accept="image/*" multiple onChange={onUpload} />
              </label>
            </div>

            <div className="flex justify-between items-center px-1">
              <span className="text-sm font-bold flex items-center gap-2" style={{ color: 'var(--m3-on-surface)' }}>
                {searchQuery ? '検索結果' : 'ライブラリ'}
                <span className="px-2 py-0.5 rounded-full text-xs" style={{ background: 'var(--m3-secondary-container)', color: 'var(--m3-on-secondary-container)' }}>{filteredImages.length}</span>
              </span>
              <button
                onClick={() => {
                  setIsImageSelectionMode(!isImageSelectionMode);
                  setSelectedImageIds(new Set());
                }}
                onMouseEnter={(e) => onShowQuickHelp?.(e, 'ライブラリのチェックボックス', '複数画像を選択するモードです。ONで一括選択・一括削除が使えます。')}
                onMouseLeave={() => onHideQuickHelp?.()}
                className={`p-2 rounded-full transition-all ${isImageSelectionMode ? 'bg-primary-container text-on-primary-container ring-1' : ''}`}
                style={isImageSelectionMode ? { background: 'var(--m3-primary-container)', color: 'var(--m3-on-primary-container)' } : { color: 'var(--m3-on-surface-variant)' }}
                title="画像を選択して削除"
              >
                <CheckSquare size={20} />
              </button>
            </div>

            {isImageSelectionMode && (
              <div className="flex items-center justify-between rounded-lg p-2.5 m3-animate-fade-in" style={{ background: 'var(--m3-surface-container-high)' }}>
                <span className="text-xs font-bold" style={{ color: 'var(--m3-primary)' }}>{selectedImageIds.size}枚選択</span>
                <div className="flex gap-2">
                  <button
                    onClick={() => setSelectedImageIds(new Set(filteredImages.filter((image) => !image.assignment).map((image) => getImageSelectionKey(image))))}
                    className="text-xs font-medium px-3 py-1.5 rounded-full hover:bg-black/5"
                    style={{ color: 'var(--m3-primary)' }}
                  >
                    全選択
                  </button>
                  <button
                    onClick={handleBulkDelete}
                    disabled={selectedImageIds.size === 0}
                    className={`text-xs px-4 py-1.5 rounded-full font-medium transition-all ${selectedImageIds.size > 0 ? 'shadow-sm' : 'opacity-50 cursor-not-allowed'}`}
                    style={{ background: 'var(--m3-error)', color: 'var(--m3-on-error)' }}
                  >
                    削除
                  </button>
                </div>
              </div>
            )}

            <div className="grid grid-cols-2 gap-2" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))' }}>
              {filteredImages.map((img) => {
                const selectionKey = getImageSelectionKey(img);
                const assignment = img.assignment || null;
                const assignmentGenre = assignment
                  ? (GENRES.find((genre) => genre.id === assignment.genre) || GENRES[0])
                  : null;
                return (
                <div
                  key={selectionKey}
                  className={`group relative rounded-lg border p-1.5 transition-all duration-200
                    ${assignment
                      ? 'cursor-pointer hover:-translate-y-0.5'
                      : isImageSelectionMode
                      ? 'cursor-pointer'
                      : 'cursor-grab active:cursor-grabbing hover:shadow-md hover:-translate-y-0.5'}
                    ${selectedImageIds.has(selectionKey)
                      ? 'ring-2'
                      : ''}`}
                  style={{
                    background: 'var(--m3-surface)',
                    borderColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'var(--m3-outline-variant)',
                    backgroundColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary-container)' : 'var(--m3-surface)',
                    touchAction: isImageSelectionMode || assignment ? 'manipulation' : 'pan-y',
                    ...(assignment ? {
                      borderColor: assignmentGenre.color,
                      boxShadow: `0 0 7px ${assignmentGenre.color}, 0 0 16px ${assignmentGenre.color}99, inset 0 0 5px ${assignmentGenre.color}55`
                    } : {})
                  }}
                  role={assignment ? 'button' : undefined}
                  tabIndex={assignment ? 0 : undefined}
                  aria-label={assignment ? `配置済み画像、ページ${assignment.sheetNumber}を開く` : undefined}
                  draggable={!isImageSelectionMode && !assignment}
                  onPointerDown={!isImageSelectionMode && !assignment ? (e) => onStartPointerDrag?.(e, {
                    payload: {
                      src: img.data,
                      imageId: img.id || '',
                      type: 'image',
                      name: img.name || '',
                      code: img.code || '',
                      freeLabels: img.freeLabels || [],
                      freeText: img.freeText || ''
                    },
                    preview: {
                      image: img.data,
                      code: img.code || img.name || ''
                    }
                  }) : undefined}
                  onClick={() => {
                    if (assignment) {
                      onOpenAssignedImage?.(assignment.sheetId);
                      return;
                    }
                    if (isImageSelectionMode) {
                      toggleImageSelection(selectionKey);
                    }
                  }}
                  onKeyDown={(e) => {
                    if (!assignment || (e.key !== 'Enter' && e.key !== ' ')) return;
                    e.preventDefault();
                    onOpenAssignedImage?.(assignment.sheetId);
                  }}
                  onDoubleClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    if (isImageSelectionMode || assignment) return;
                    handleTogglePreview(img.data, img.name || '');
                  }}
                  onDragStart={(e) => {
                    if (isImageSelectionMode || assignment) {
                      e.preventDefault();
                      return;
                    }
                    setDragPayload(e.dataTransfer, {
                      src: img.data,
                      imageId: img.id || '',
                      type: 'image',
                      name: img.name || '',
                      code: img.code || '',
                      freeLabels: img.freeLabels || [],
                      freeText: img.freeText || ''
                    });
                  }}
                >
                  <div className="relative mb-1 aspect-square w-full overflow-hidden rounded-md bg-white flex items-center justify-center">
                    <img src={img.data} alt="stock" className="max-w-full max-h-full object-contain" loading="lazy" decoding="async" draggable={false} />
                    <FreeLabelPreview item={img} />
                    {assignment && (
                      <div className="pointer-events-none absolute inset-0 z-20 flex flex-col items-center justify-center bg-white/10 text-center">
                        <span
                          className="rounded-lg bg-white/65 px-2.5 py-1 text-lg font-black tracking-tight shadow-sm backdrop-blur-[1px]"
                          style={{ color: assignmentGenre.color, textShadow: '0 1px 2px rgba(15, 23, 42, 0.25)' }}
                        >
                          Page {assignment.sheetNumber}
                        </span>
                        <span className="mt-1 rounded bg-slate-900/45 px-1.5 py-0.5 text-[9px] font-bold text-white">
                          {assignmentGenre.label}
                        </span>
                      </div>
                    )}
                  </div>
                  <div className="px-1">
                    <div className="text-[10px] truncate font-medium" style={{ color: 'var(--m3-on-surface)' }}>{img.name}</div>
                  </div>

                  {isImageSelectionMode && !assignment ? (
                    <div className="absolute top-2 left-2 w-6 h-6 rounded-full flex items-center justify-center transition-all"
                      style={{ background: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'rgba(255,255,255,0.8)', border: selectedImageIds.has(selectionKey) ? 'none' : '1px solid var(--m3-outline)' }}>
                      {selectedImageIds.has(selectionKey) && <Check size={14} style={{ color: 'var(--m3-on-primary)' }} />}
                    </div>
                  ) : !assignment ? (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteImage(img.id, img.data);
                      }}
                      className="absolute top-2 right-2 rounded-full p-2 opacity-0 group-hover:opacity-100 shadow-sm transition-all hover:scale-110"
                      style={{ background: 'var(--m3-surface)', color: 'var(--m3-error)', border: '1px solid var(--m3-outline-variant)' }}
                    >
                      <Trash2 size={14} />
                    </button>
                  ) : null}
                </div>
                );
              })}

              {filteredImages.length === 0 && (
                <div className="col-span-full flex flex-col items-center justify-center py-16 text-center" style={{ color: 'var(--m3-outline)' }}>
                  <Search size={32} className="mb-3 opacity-20" />
                  <p className="text-sm">{searchQuery ? '見つかりませんでした' : '画像がありません'}</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Dummy Tab */}
        {activeTab === 'dummy' && (
          <div className="space-y-2.5">
            <div className="flex items-center justify-between px-0.5">
              <p className="text-xs font-bold" style={{ color: 'var(--m3-on-surface)' }}>ダミーコマ</p>
              <p className="text-[10px]" style={{ color: 'var(--m3-outline)' }}>ドラッグして配置</p>
            </div>

            {DUMMY_OPTIONS.map((dummy) => (
              <div
                key={dummy.label}
                className="flex h-16 cursor-grab items-center justify-between rounded-lg border px-4 active:cursor-grabbing hover:-translate-y-0.5 hover:shadow-sm transition-all"
                style={{ background: dummy.color, borderColor: 'rgba(100, 116, 139, 0.24)', color: dummy.text, touchAction: 'pan-y' }}
                draggable
                onPointerDown={(e) => {
                  const payload = createDummyDragPayload(dummy);
                  onStartPointerDrag?.(e, {
                    payload,
                    preview: {
                      image: payload.src,
                      label: dummy.label,
                      text: dummy.isText ? 'テキスト' : ''
                    }
                  });
                }}
                onDragStart={(e) => {
                  setDragPayload(e.dataTransfer, createDummyDragPayload(dummy));
                }}
              >
                <span className="text-sm font-bold">{dummy.label}</span>
                <span className="rounded-full bg-white/55 px-2 py-1 text-[9px] font-bold">配置</span>
              </div>
            ))}
          </div>
        )}


        {/* Status Tab */}
        {activeTab === 'status' && (
          <div className="space-y-3">
            <div className="rounded-xl border border-slate-200 bg-white p-3">
              <div className="flex items-center justify-between">
                <div>
                  <p className="flex items-center gap-1.5 text-xs font-bold text-slate-700">
                    <Info size={14} /> 空き状況
                  </p>
                  <p className="mt-1 text-[10px] text-slate-400">{statusFilteredSheets.length}ページを表示</p>
                </div>
                <span
                  className={`rounded-full px-2.5 py-1 text-xs font-bold tabular-nums ${statusFilteredEmptyCount > 0 ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600'}`}
                  title={statusGenreFilter === 'all' ? '全ジャンルの空きコマ合計' : 'このジャンルの空きコマ合計'}
                >
                  {statusFilteredEmptyCount > 0 ? `空き ${statusFilteredEmptyCount}` : 'すべて完了'}
                </span>
              </div>
              <select
                value={statusGenreFilter}
                onChange={(e) => setStatusGenreFilter(e.target.value)}
                className="mt-2.5 w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-2 text-xs font-medium text-slate-600 focus:outline-none focus:ring-1 focus:ring-indigo-500"
              >
                <option value="all">全ジャンル</option>
                {GENRES.map(g => (
                  <option key={g.id} value={g.id}>{g.label}</option>
                ))}
              </select>
            </div>

            <div className="space-y-1.5">
              {statusFilteredSheets.map(({ sheet, originalIndex }) => {
                  const totalEmpty = getSheetEmptyCount(sheet);
                  const genre = GENRES.find(g => g.id === sheet.genre) || GENRES[0];

                  return (
                    <div
                      key={sheet.id}
                      className="flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2.5"
                      style={{ borderLeft: `4px solid ${genre.color}` }}
                    >
                      <div className="min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-bold text-slate-700">P{originalIndex + 1}</span>
                          <span className="truncate text-[10px] font-medium text-slate-500">{genre.label}</span>
                        </div>
                      </div>
                      <span className={`font-mono text-xs font-bold ${totalEmpty > 0 ? 'text-rose-500' : 'text-emerald-500'}`}>
                        {totalEmpty > 0 ? `${totalEmpty}コマ` : '完了'}
                      </span>
                    </div>
                  );
                })}
              {statusFilteredSheets.length === 0 && (
                <div className="rounded-lg border border-dashed border-slate-200 py-8 text-center text-xs text-slate-400">
                  対象ページがありません
                </div>
              )}
            </div>
          </div>
        )}
        {activeTab === 'excluded' && (
          <div
            className="flex h-full flex-1 flex-col overflow-hidden rounded-xl border border-slate-200 bg-white"
            data-daiwari-dropzone-id="excluded"
            onDragOverCapture={(e) => e.preventDefault()}
            onDropCapture={handleDropToExcluded}
            onDragOver={(e) => e.preventDefault()}
            onDrop={handleDropToExcluded}
          >
            <div className="space-y-2.5 border-b border-slate-200 p-3">
              <div className="flex items-center justify-between">
                <div className="flex min-w-0 items-center gap-2">
                  <Ban size={15} className="flex-shrink-0 text-rose-500" />
                  <span className="truncate text-xs font-bold text-slate-700">掲載除外</span>
                  <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[10px] font-bold text-slate-500">
                    {excludedSearchQuery ? `${filteredExcludedItems.length}/${excludedItems.length}` : excludedItems.length}件
                  </span>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={onExportExcludedCSV}
                    disabled={excludedItems.length === 0}
                    className="flex items-center gap-1 rounded-md border border-slate-200 px-2 py-1 text-[10px] font-bold text-slate-600 transition-colors hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-35"
                    title="除外リストをCSVで出力"
                  >
                    <FileSpreadsheet size={12} /> CSV
                  </button>
                  <button
                    onClick={onBulkDeleteExcluded}
                    disabled={excludedItems.length === 0}
                    className="rounded-md border border-slate-200 p-1.5 text-rose-500 transition-colors hover:border-rose-200 hover:bg-rose-50 disabled:cursor-not-allowed disabled:opacity-35"
                    title="除外リストを全て空にする"
                    aria-label="除外リストを全て空にする"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              </div>
              <div className="relative">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="コード・ラベル・名前で検索"
                  value={excludedSearchQuery}
                  onChange={(e) => setExcludedSearchQuery(e.target.value)}
                  onMouseEnter={(e) => onShowQuickHelp?.(e, '除外リスト検索', 'コード・ラベル・画像名で除外リストを絞り込みます。')}
                  onMouseLeave={() => onHideQuickHelp?.()}
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2 pl-9 pr-8 text-xs text-slate-700 placeholder-slate-400 transition-colors focus:border-indigo-300 focus:outline-none focus:ring-1 focus:ring-indigo-200"
                />
                {excludedSearchQuery && (
                  <button
                    type="button"
                    onClick={() => setExcludedSearchQuery('')}
                    className="absolute right-2 top-1/2 -translate-y-1/2 rounded-full p-0.5 text-slate-400 transition-colors hover:bg-slate-200 hover:text-slate-600"
                    title="検索クリア"
                  >
                    <X size={12} />
                  </button>
                )}
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-2.5">
              {excludedItems.length === 0 ? (
                <div className="m-1 flex h-full flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 py-8 text-slate-400">
                  <Ban size={22} className="mb-2 opacity-50" />
                  <p className="text-xs font-medium">ここにドロップして除外</p>
                </div>
              ) : filteredExcludedItems.length === 0 ? (
                <div className="m-1 flex h-full flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 py-8 text-slate-400">
                  <Search size={22} className="mb-2 opacity-50" />
                  <p className="text-xs font-medium">検索結果なし</p>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  {filteredExcludedItems.map((item) => {
                    const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                    return (
                      <div
                        key={item.id}
                        className="group relative flex cursor-grab flex-col items-center rounded-lg border border-slate-200 bg-white p-1.5 transition-all hover:border-slate-300 hover:shadow-sm active:cursor-grabbing"
                        style={{ touchAction: 'none' }}
                        draggable
                        onPointerDown={(e) => {
                          const payloadText = typeof item.text === 'string' ? item.text : '';
                          onStartPointerDrag?.(e, {
                            payload: {
                              src: resolvedImg || '',
                              type: 'image',
                              name: item.originalName || 'excluded',
                              label: item.label || '',
                              code: item.code || '',
                              isText: item.isText ? 'true' : 'false',
                              hasTextPayload: '1',
                              textPayload: payloadText,
                              text: payloadText,
                              freeLabels: item.freeLabels || [],
                              freeText: item.freeText || '',
                              fromExcludedId: item.id,
                              imageId: item.imageId || ''
                            },
                            preview: {
                              image: resolvedImg || null,
                              label: item.label || null,
                              code: item.code || null,
                              text: item.isText ? payloadText : ''
                            }
                          });
                        }}
                        onDoubleClick={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          handleTogglePreview(resolvedImg || '', item.code || item.originalName || '');
                        }}
                        onDragStart={(e) => {
                          const payloadText = typeof item.text === 'string' ? item.text : '';
                          setDragPayload(e.dataTransfer, {
                            src: resolvedImg || '',
                            type: 'image',
                            name: item.originalName || 'excluded',
                            label: item.label || '',
                            code: item.code || '',
                            isText: item.isText ? 'true' : 'false',
                            hasTextPayload: '1',
                            textPayload: payloadText,
                            text: payloadText,
                            freeLabels: item.freeLabels || [],
                            freeText: item.freeText || '',
                            fromExcludedId: item.id,
                            imageId: item.imageId || ''
                          });
                        }}
                      >
                        <div className="relative mb-1 w-full aspect-square overflow-hidden rounded-md bg-slate-50">
                          {resolvedImg ? (
                            <img src={resolvedImg} alt="excluded" className="w-full h-full object-contain" draggable={false} />
                          ) : (
                            <div className="w-full h-full flex flex-col items-center justify-center text-slate-300">
                              <Ban size={16} />
                            </div>
                          )}
                          <FreeLabelPreview item={item} />
                        </div>

                        <div className="flex w-full min-w-0 items-center gap-1 text-[9px]">
                          <span className="min-w-0 flex-1 truncate font-mono font-bold text-slate-700">{item.code || item.originalName || 'コードなし'}</span>
                          {item.label && (
                            <span className="max-w-[48px] truncate rounded bg-slate-100 px-1 text-slate-500">{item.label}</span>
                          )}
                        </div>

                        <button
                          onClick={(event) => {
                            event.stopPropagation();
                            onDeleteFromExcluded(item.id);
                          }}
                          onPointerDown={(event) => event.stopPropagation()}
                          className="absolute right-1 top-1 rounded-full border border-slate-200 bg-white/90 p-1 text-slate-400 shadow-sm transition-colors hover:border-rose-200 hover:bg-rose-50 hover:text-rose-600"
                          title="完全に削除"
                        >
                          <X size={12} strokeWidth={3} />
                        </button>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Temp Shelf (Fixed at bottom) */}
      <div
        className="relative flex h-52 flex-shrink-0 flex-col border-t border-slate-200 bg-slate-50"
        data-daiwari-dropzone-id="temp"
        onDragOverCapture={(e) => e.preventDefault()}
        onDropCapture={handleDropToTemp}
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDropToTemp}
      >
        <div
          className="flex items-center justify-between border-b border-slate-200 bg-white px-3 py-2"
          onMouseEnter={(e) => onShowQuickHelp?.(e, '仮置き場', 'コマを一時退避する場所です。ログイン中のGoogleアカウント専用の仮置き場です。')}
          onMouseLeave={() => onHideQuickHelp?.()}
        >
          <div className="flex min-w-0 items-center gap-1.5 text-xs font-bold text-slate-700">
            <ClipboardList size={14} className="flex-shrink-0 text-indigo-500" />
            <span className="truncate">仮置き場</span>
            <span className="rounded-full bg-indigo-50 px-1.5 py-0.5 text-[10px] text-indigo-600">
              {activeTempItems.length}件
            </span>
          </div>
          <span className="text-[10px] font-medium text-slate-400">自分専用</span>
        </div>

        <div className="flex-1 overflow-y-auto p-2">
          {activeTempItems.length === 0 ? (
            <div className="flex h-full flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 text-slate-400">
              <ClipboardList size={18} className="mb-1 opacity-50" />
              <p className="text-[11px] font-medium">ここにドロップ</p>
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-2">
              {activeTempItems.map((item) => {
                const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                const hoverCodeText = (item.code || '').trim();
                return (
                  <div
                    key={item.id}
                    className="group relative flex min-h-[78px] cursor-grab flex-col items-center rounded-lg border border-slate-200 bg-white p-1.5 transition-all hover:border-slate-300 hover:shadow-sm active:cursor-grabbing"
                    title={hoverCodeText || undefined}
                    style={{ touchAction: 'none' }}
                    draggable
                    onPointerDown={(e) => {
                      const payloadText = typeof item.text === 'string' ? item.text : '';
                      onStartPointerDrag?.(e, {
                        payload: {
                          src: resolvedImg || '',
                          type: 'image',
                          name: item.originalName || 'temp',
                          label: item.label || '',
                          code: item.code || '',
                          isText: item.isText ? 'true' : 'false',
                          hasTextPayload: '1',
                          textPayload: payloadText,
                          text: payloadText,
                          freeLabels: item.freeLabels || [],
                          freeText: item.freeText || '',
                          fromTempId: item.id,
                          imageId: item.imageId || ''
                        },
                        preview: {
                          image: resolvedImg || null,
                          label: item.label || null,
                          code: item.code || null,
                          text: item.isText ? payloadText : ''
                        }
                      });
                    }}
                    onDoubleClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      handleTogglePreview(resolvedImg || '', hoverCodeText || item.originalName || '');
                    }}
                    onDragStart={(e) => {
                      const payloadText = typeof item.text === 'string' ? item.text : '';
                      setDragPayload(e.dataTransfer, {
                        src: resolvedImg || '',
                        type: 'image',
                        name: item.originalName || 'temp',
                        label: item.label || '',
                        code: item.code || '',
                        isText: item.isText ? 'true' : 'false',
                        hasTextPayload: '1',
                        textPayload: payloadText,
                        text: payloadText,
                        freeLabels: item.freeLabels || [],
                        freeText: item.freeText || '',
                        fromTempId: item.id,
                        imageId: item.imageId || ''
                      });
                    }}
                  >
                    <div className="relative h-14 w-full overflow-hidden rounded-md bg-slate-50">
                      {resolvedImg ? (
                        <img src={resolvedImg} alt="temp" className="w-full h-full object-contain" draggable={false} />
                      ) : (
                        <div className="w-full h-full flex flex-col items-center justify-center bg-slate-50 text-slate-400">
                          <span className="text-[10px] font-mono">{item.code || 'No Image'}</span>
                        </div>
                      )}
                      <FreeLabelPreview item={item} />
                    </div>

                    {item.label && (
                      <div className="absolute left-1 top-1 max-w-[65%] truncate rounded bg-slate-800/75 px-1.5 py-0.5 text-[8px] text-white">
                        {item.label}
                      </div>
                    )}
                    <p className="mt-1 w-full truncate text-center font-mono text-[9px] font-bold text-slate-600">
                      {hoverCodeText || item.originalName || '仮置き'}
                    </p>
                    <button
                      onClick={(event) => {
                        event.stopPropagation();
                        onDeleteFromTemp(item.id);
                      }}
                      onPointerDown={(event) => event.stopPropagation()}
                      className="absolute right-1 top-1 rounded-full border border-slate-200 bg-white/90 p-1 text-slate-400 shadow-sm transition-colors hover:border-rose-200 hover:bg-rose-50 hover:text-rose-600"
                      title="仮置き場から削除"
                    >
                      <X size={12} strokeWidth={3} />
                    </button>
                  </div>
                )
              })}
            </div>
          )}
        </div>

      </div>

      <ImagePreviewModal preview={previewImage} onClose={() => setPreviewImage(null)} />

      <div
        className="absolute right-0 top-0 bottom-0 w-1 cursor-ew-resize hover:bg-indigo-400 transition-colors z-30"
        onMouseDown={() => setResizing(true)}
      />
    </div>
  );
});


export default Sidebar;
