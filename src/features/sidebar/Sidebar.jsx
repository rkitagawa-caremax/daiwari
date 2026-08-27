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

// ... Sidebar Component ...
const Sidebar = React.memo(({
  isLocked,
  isOpen,
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

  // 画像検索フィルタリング
  // - 配置済み画像 (どこかのシートのコマに使われているもの) は除外
  // - 掲載除外リスト (excludedItems) に入っている画像も除外 (除外タブには引き続き表示)
  const filteredImages = useMemo(() => {
    // 全シートで使用されている画像のSetを作成
    const usedImageSet = new Set();
    sheets.forEach(sheet => {
      sheet.panels.forEach(p => {
        if (p.image) usedImageSet.add(p.image);
        if (p.imageId) usedImageSet.add(p.imageId);
      });
    });

    // 掲載除外リスト (excludedItems) に入っている画像のキーセット
    const excludedImageSet = new Set();
    (excludedItems || []).forEach((item) => {
      if (!item) return;
      if (item.image) excludedImageSet.add(item.image);
      if (item.imageId) excludedImageSet.add(item.imageId);
    });

    let result = images.filter(img =>
      !usedImageSet.has(img.data)
      && !usedImageSet.has(img.id)
      && !excludedImageSet.has(img.data)
      && !excludedImageSet.has(img.id)
    );

    if (!searchQuery) return result;
    const lowerQuery = searchQuery.toLowerCase();
    return result.filter(img => (img.name || '').toLowerCase().includes(lowerQuery));
  }, [images, searchQuery, sheets, excludedItems]);

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

  // 空き状況タブのジャンル絞り込み後の合計空きコマ数。
  // statusGenreFilter === 'all' なら全シート、ジャンル指定なら該当ジャンルのみ。
  // 計算ロジックは下のシート一覧 (pureEmptyCount + dummyCount) と完全に一致する。
  const statusFilteredEmptyCount = useMemo(() => {
    return (sheets || [])
      .filter((sheet) => statusGenreFilter === 'all' || sheet.genre === statusGenreFilter)
      .reduce((total, sheet) => {
        const panels = Array.isArray(sheet.panels) ? sheet.panels : [];
        const pureEmpty = panels.filter((p) => !p.hidden && !p.image && !p.imageId && !p.label).length;
        const dummy = panels.filter((p) => !p.hidden && p.label && p.label !== 'タイトル' && p.label !== '埋草' && p.label !== 'テキスト').length;
        return total + pureEmpty + dummy;
      }, 0);
  }, [sheets, statusGenreFilter]);

  if (!isOpen) {
    return (
      <div className="fixed left-0 top-16 bottom-0 w-16 m3-surface border-r flex flex-col items-center py-6 z-20 transition-all duration-300" style={{ borderColor: 'var(--m3-outline-variant)' }}>
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
      className="fixed left-0 top-16 bottom-0 m3-surface border-r flex flex-col z-20 shadow-xl transition-all duration-300 ease-in-out"
      style={{ width, borderColor: 'var(--m3-outline-variant)' }}
    >
      <div className="flex items-center justify-between px-4 py-4 border-b flex-shrink-0" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface-container)' }}>
        <div className="flex p-1 rounded-full flex-1 mr-4 overflow-x-auto" style={{ background: 'var(--m3-surface-container-highest)' }}>
          {[
            { id: 'stock', icon: ImageIcon, label: '画像' },
            { id: 'dummy', icon: Type, label: 'ダミー' },
            { id: 'status', icon: Info, label: '空き' },
            { id: 'excluded', icon: Ban, label: '除外' }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              onMouseEnter={(e) => {
                const help = sidebarTabHelp[tab.id];
                if (help) onShowQuickHelp?.(e, help.title, help.description);
              }}
              onMouseLeave={() => onHideQuickHelp?.()}
              className={`flex-1 flex items-center justify-center gap-1.5 py-1.5 px-2 rounded-full transition-all duration-200 whitespace-nowrap ${activeTab === tab.id ? 'bg-white shadow-sm' : 'hover:bg-black/5 opacity-70'}`}
              style={activeTab === tab.id ? { color: 'var(--m3-primary)' } : { color: 'var(--m3-on-surface-variant)' }}
            >
              <tab.icon size={16} />
              <span className="text-xs font-medium">{tab.label}</span>
            </button>
          ))}
        </div>
        <button onClick={toggleOpen} className="m3-icon-btn flex-shrink-0">
          <ChevronLeft size={24} />
        </button>
      </div>

      {activeTab === 'stock' && (
        <div className="p-4 border-b flex-shrink-0" style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}>
          <div className="relative group">
            <Search className="absolute left-4 top-3.5 w-5 h-5 transition-colors" style={{ color: 'var(--m3-outline)' }} />
            <input
              type="text"
              placeholder="画像検索..."
              value={searchQuery}
              onChange={(e) => onSearch(e.target.value)}
              onMouseEnter={(e) => onShowQuickHelp?.(e, '画像検索', '画像名でライブラリを絞り込みます。キーワード入力で候補を素早く探せます。')}
              onMouseLeave={() => onHideQuickHelp?.()}
              className="w-full p-3 pl-12 rounded-full transition-all"
              style={{ background: 'var(--m3-surface-container-high)', color: 'var(--m3-on-surface)' }}
            />
            {searchQuery && (
              <button
                onClick={() => onSearch('')}
                className="absolute right-3 top-3 p-1 rounded-full hover:bg-black/10 transition-colors"
                style={{ color: 'var(--m3-on-surface-variant)' }}
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>
      )}

      <div
        className="flex-1 overflow-y-auto p-4 relative"
        style={{ background: 'var(--m3-surface-container-low)' }}
        data-daiwari-dropzone-id={activeTab === 'stock' ? 'stock' : undefined}
        onDragOverCapture={(e) => e.preventDefault()}
        onDropCapture={activeTab === 'stock' ? handleDropToStock : undefined}
        onDragOver={activeTab === 'stock' ? (e) => e.preventDefault() : undefined}
        onDrop={activeTab === 'stock' ? handleDropToStock : undefined}
      >

        {activeTab === 'stock' && (
          <div className="space-y-6">
            <div className="space-y-2">
              <label
                className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl cursor-pointer transition-all group"
                style={{ borderColor: 'var(--m3-outline-variant)', background: 'var(--m3-surface)' }}
                onMouseEnter={(e) => onShowQuickHelp?.(e, 'クリックしてアップロード', '画像ファイルを選択してライブラリに追加します。複数選択アップロードにも対応します。')}
                onMouseLeave={() => onHideQuickHelp?.()}
              >
                <div className="flex flex-col items-center pt-2 pb-3">
                  <div className="p-3 rounded-full mb-3 transition-colors" style={{ background: 'var(--m3-primary-container)' }}>
                    <ImageIcon className="w-6 h-6" style={{ color: 'var(--m3-on-primary-container)' }} />
                  </div>
                  <p className="text-sm font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>クリックしてアップロード</p>
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
              <div className="flex items-center justify-between p-4 rounded-xl shadow-sm m3-animate-fade-in" style={{ background: 'var(--m3-surface-container-high)' }}>
                <span className="text-sm font-bold" style={{ color: 'var(--m3-primary)' }}>{selectedImageIds.size}枚選択中</span>
                <div className="flex gap-2">
                  <button
                    onClick={() => setSelectedImageIds(new Set(filteredImages.map((i) => getImageSelectionKey(i))))}
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

            <div className="grid grid-cols-2 gap-3" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))' }}>
              {filteredImages.map((img) => {
                const selectionKey = getImageSelectionKey(img);
                return (
                <div
                  key={selectionKey}
                  className={`group relative border rounded-xl p-2 transition-all duration-200
                    ${isImageSelectionMode
                      ? 'cursor-pointer'
                      : 'cursor-grab active:cursor-grabbing hover:shadow-md hover:-translate-y-0.5'}
                    ${selectedImageIds.has(selectionKey)
                      ? 'ring-2'
                      : ''}`}
                  style={{
                    background: 'var(--m3-surface)',
                    borderColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'var(--m3-outline-variant)',
                    backgroundColor: selectedImageIds.has(selectionKey) ? 'var(--m3-primary-container)' : 'var(--m3-surface)',
                    touchAction: isImageSelectionMode ? 'manipulation' : 'pan-y'
                  }}
                  draggable={!isImageSelectionMode}
                  onPointerDown={!isImageSelectionMode ? (e) => onStartPointerDrag?.(e, {
                    payload: {
                      src: img.data,
                      imageId: img.id || '',
                      type: 'image',
                      name: img.name || ''
                    },
                    preview: {
                      image: img.data,
                      code: img.name || ''
                    }
                  }) : undefined}
                  onClick={() => {
                    if (isImageSelectionMode) {
                      toggleImageSelection(selectionKey);
                    }
                  }}
                  onDoubleClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    if (isImageSelectionMode) return;
                    handleTogglePreview(img.data, img.name || '');
                  }}
                  onDragStart={(e) => {
                    if (isImageSelectionMode) {
                      e.preventDefault();
                      return;
                    }
                    setDragPayload(e.dataTransfer, {
                      src: img.data,
                      imageId: img.id || '',
                      type: 'image',
                      name: img.name || ''
                    });
                  }}
                >
                  <div className="aspect-square w-full rounded-lg overflow-hidden mb-2 bg-white flex items-center justify-center">
                    <img src={img.data} alt="stock" className="max-w-full max-h-full object-contain" loading="lazy" decoding="async" draggable={false} />
                  </div>
                  <div className="px-1">
                    <div className="text-[10px] truncate font-medium" style={{ color: 'var(--m3-on-surface)' }}>{img.name}</div>
                  </div>

                  {isImageSelectionMode ? (
                    <div className="absolute top-2 left-2 w-6 h-6 rounded-full flex items-center justify-center transition-all"
                      style={{ background: selectedImageIds.has(selectionKey) ? 'var(--m3-primary)' : 'rgba(255,255,255,0.8)', border: selectedImageIds.has(selectionKey) ? 'none' : '1px solid var(--m3-outline)' }}>
                      {selectedImageIds.has(selectionKey) && <Check size={14} style={{ color: 'var(--m3-on-primary)' }} />}
                    </div>
                  ) : (
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
                  )}
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
          <div className="space-y-4">
            <p className="text-xs font-medium" style={{ color: 'var(--m3-on-surface-variant)' }}>ドラッグして配置できます</p>

            {[
              { color: 'var(--dummy-gray)', text: 'var(--dummy-text-color)', label: '新規商品未確定' },
              { color: 'var(--dummy-red)', text: 'var(--dummy-text-color)', label: 'タイトル' },
              { color: 'var(--dummy-green)', text: 'var(--dummy-text-color)', label: '埋草' },
              { color: 'var(--dummy-volt)', text: 'var(--dummy-text-color)', label: 'テキスト', isText: true },
            ].map((dummy) => (
              <div
                key={dummy.label}
                className="h-24 border-2 border-dashed rounded-xl flex items-center justify-center cursor-grab active:cursor-grabbing hover:shadow-md hover:-translate-y-0.5 transition-all"
                style={{ background: dummy.color, borderColor: 'var(--m3-outline-variant)', color: dummy.text, touchAction: 'pan-y' }}
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
                <span className="font-bold text-sm">【{dummy.label}】</span>
              </div>
            ))}
          </div>
        )}


        {/* Status Tab */}
        {activeTab === 'status' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="font-bold text-slate-600 flex items-center gap-2 text-xs">
                <Info size={16} /> 空き状況
              </h3>
              <div className="flex items-center gap-2">
                <span
                  className="text-xs font-bold text-rose-500 tabular-nums"
                  title={statusGenreFilter === 'all' ? '全ジャンルの空きコマ合計' : 'このジャンルの空きコマ合計'}
                >
                  空き {statusFilteredEmptyCount}
                </span>
                <select
                  value={statusGenreFilter}
                  onChange={(e) => setStatusGenreFilter(e.target.value)}
                  className="text-xs border border-slate-200 rounded-lg px-2 py-1.5 bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 text-slate-600"
                >
                  <option value="all">全ジャンル</option>
                  {GENRES.map(g => (
                    <option key={g.id} value={g.id}>{g.label}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="space-y-2">
              {sheets
                .map((sheet, originalIndex) => ({ sheet, originalIndex }))
                .filter(({ sheet }) => statusGenreFilter === 'all' || sheet.genre === statusGenreFilter)
                .map(({ sheet, originalIndex }) => {
                  const pureEmptyCount = sheet.panels.filter(p => !p.hidden && !p.image && !p.imageId && !p.label).length;
                  const dummyCount = sheet.panels.filter(p => !p.hidden && p.label && p.label !== 'タイトル' && p.label !== '埋草' && p.label !== 'テキスト').length;
                  const totalEmpty = pureEmptyCount + dummyCount;

                  const dummyDetails = sheet.panels.reduce((acc, p) => {
                    if (!p.hidden && p.label && p.label !== 'タイトル' && p.label !== '埋草') {
                      acc[p.label] = (acc[p.label] || 0) + 1;
                    }
                    return acc;
                  }, {});

                  const genre = GENRES.find(g => g.id === sheet.genre) || GENRES[0];

                  return (
                    <div key={sheet.id} className="flex flex-col bg-white border border-slate-100 rounded-lg p-3 shadow-sm hover:shadow-md transition-shadow">
                      <div className="flex justify-between items-center mb-2">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-slate-700 text-xs">P{originalIndex + 1}</span>
                          <span
                            className="text-[10px] px-2 py-0.5 rounded-full font-medium truncate max-w-[100px]"
                            style={{
                              backgroundColor: genre.color,
                              color: '#1e293b' // Darker text for better contrast
                            }}
                          >
                            {genre.label}
                          </span>
                        </div>
                        <span className={`font-mono text-xs font-bold ${totalEmpty > 0 ? 'text-rose-500' : 'text-emerald-500'}`}>
                          {totalEmpty > 0 ? `空き: ${totalEmpty}` : '完了'}
                        </span>
                      </div>

                      {Object.keys(dummyDetails).length > 0 && (
                        <div className="flex flex-wrap gap-1.5 pl-4 border-l-2 border-slate-200 ml-1">
                          {Object.entries(dummyDetails).map(([label, count]) => {
                            const dummyAttr = [
                              { color: 'bg-[var(--dummy-gray)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: '新規商品未確定' },
                              { color: 'bg-[var(--dummy-orange)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'その他' },
                              { color: 'bg-[var(--dummy-red)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'タイトル' },
                              { color: 'bg-[var(--dummy-green)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: '埋草' },
                              { color: 'bg-[var(--dummy-volt)]', border: 'border-transparent', text: 'text-[var(--dummy-text-color)]', label: 'テキスト' },
                            ].find(d => d.label === label) || { color: 'bg-slate-50', text: 'text-slate-500', border: 'border-slate-200' };

                            return (
                              <span key={label} className={`text-[9.5px] px-2 py-0.5 ${dummyAttr.color} ${dummyAttr.text} rounded-full border ${dummyAttr.border} font-bold shadow-sm`}>
                                {label}: {count}
                              </span>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              {sheets.length === 0 && <p className="text-xs text-slate-400 text-center py-4">ページがありません</p>}
            </div>
          </div>
        )}
        {activeTab === 'excluded' && (
          <div
            className="flex-1 flex flex-col h-full bg-rose-50/30 rounded-xl border border-rose-100 overflow-hidden"
            data-daiwari-dropzone-id="excluded"
            onDragOverCapture={(e) => e.preventDefault()}
            onDropCapture={handleDropToExcluded}
            onDragOver={(e) => e.preventDefault()}
            onDrop={handleDropToExcluded}
          >
            <div className="p-2 border-b border-rose-100 bg-white/60">
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-rose-300" />
                <input
                  type="text"
                  placeholder="除外リストを検索 (コード・ラベル・名前)"
                  value={excludedSearchQuery}
                  onChange={(e) => setExcludedSearchQuery(e.target.value)}
                  onMouseEnter={(e) => onShowQuickHelp?.(e, '除外リスト検索', 'コード・ラベル・画像名で除外リストを絞り込みます。')}
                  onMouseLeave={() => onHideQuickHelp?.()}
                  className="w-full pl-9 pr-8 py-1.5 text-xs rounded-full bg-white border border-rose-200 text-rose-700 placeholder-rose-300 focus:outline-none focus:border-rose-400 focus:ring-1 focus:ring-rose-200 transition-colors"
                />
                {excludedSearchQuery && (
                  <button
                    type="button"
                    onClick={() => setExcludedSearchQuery('')}
                    className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 rounded-full text-rose-400 hover:bg-rose-50 hover:text-rose-600 transition-colors"
                    title="検索クリア"
                  >
                    <X size={12} />
                  </button>
                )}
              </div>
            </div>
            <div className="p-3 border-b border-rose-100 flex items-center justify-between bg-white/50">
              <div className="flex items-center gap-2 text-sm font-bold text-rose-600">
                <Ban size={16} />
                <span>掲載除外リスト</span>
                {excludedSearchQuery && excludedItems.length > 0 && (
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-rose-50 text-rose-500 border border-rose-100 font-medium">
                    {filteredExcludedItems.length}/{excludedItems.length}件
                  </span>
                )}
              </div>
              <div className="flex items-center gap-1">
                <button
                  onClick={onExportExcludedCSV}
                  disabled={excludedItems.length === 0}
                  className={`flex items-center gap-1.5 px-2 py-1 text-xs rounded-md border transition-colors font-medium ${excludedItems.length > 0 ? 'bg-white text-rose-600 border-rose-200 hover:bg-rose-50' : 'bg-slate-50 text-slate-300 border-transparent cursor-not-allowed'}`}
                  title="除外リストをCSVで出力"
                >
                  <FileSpreadsheet size={12} />
                </button>
                <button
                  onClick={onBulkDeleteExcluded}
                  disabled={excludedItems.length === 0}
                  className={`flex items-center gap-1.5 px-2 py-1 text-xs rounded-md border transition-colors font-medium ${excludedItems.length > 0 ? 'bg-white text-rose-600 border-rose-200 hover:bg-rose-500 hover:text-white' : 'bg-slate-50 text-slate-300 border-transparent cursor-not-allowed'}`}
                  title="除外リストを全て空にする"
                >
                  <Trash2 size={12} />
                </button>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-2">
              {excludedItems.length === 0 ? (
                <div className="h-full flex flex-col items-center justify-center text-rose-300 py-8 border-2 border-dashed border-rose-200/50 rounded-lg m-1">
                  <Ban size={24} className="mb-2 opacity-50" />
                  <p className="text-[10px]">アイテムがありません</p>
                  <p className="text-[9px] opacity-70">ここへドロップして除外</p>
                </div>
              ) : filteredExcludedItems.length === 0 ? (
                <div className="h-full flex flex-col items-center justify-center text-rose-300 py-8 border-2 border-dashed border-rose-200/50 rounded-lg m-1">
                  <Search size={24} className="mb-2 opacity-50" />
                  <p className="text-[10px]">検索結果なし</p>
                  <p className="text-[9px] opacity-70">キーワードを変えて再検索</p>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  {filteredExcludedItems.map((item) => {
                    const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                    return (
                      <div
                        key={item.id}
                        className="group relative border border-rose-100 rounded-lg p-2 bg-white hover:shadow-md cursor-grab active:cursor-grabbing flex flex-col items-center transition-all"
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
                        <div className="relative w-full aspect-square bg-slate-50 rounded mb-2 overflow-hidden">
                          {resolvedImg ? (
                            <img src={resolvedImg} alt="excluded" className="w-full h-full object-contain" draggable={false} />
                          ) : (
                            <div className="w-full h-full flex flex-col items-center justify-center text-slate-300">
                              <Ban size={16} />
                            </div>
                          )}
                          <FreeLabelPreview item={item} />
                        </div>

                        <div className="w-full flex justify-between items-center text-[9px]">
                          <span className="font-bold text-rose-500 truncate max-w-[60px] font-mono">{item.code || 'No Code'}</span>
                          {item.label && (
                            <span className="bg-slate-100 text-slate-500 px-1 rounded truncate max-w-[50px]">{item.label}</span>
                          )}
                        </div>

                        <button
                          onClick={() => onDeleteFromExcluded(item.id)}
                          className="absolute -top-1.5 -right-1.5 bg-white rounded-full p-1 shadow-sm text-rose-400 hover:text-rose-600 hover:bg-rose-50 border border-rose-100 opacity-0 group-hover:opacity-100 transition-all"
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
        className="relative h-64 border-t border-slate-200 bg-slate-50 flex flex-col flex-shrink-0"
        data-daiwari-dropzone-id="temp"
        onDragOverCapture={(e) => e.preventDefault()}
        onDropCapture={handleDropToTemp}
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDropToTemp}
      >
        <div
          className="px-4 py-2 border-b border-slate-200 flex items-center justify-between bg-white"
          onMouseEnter={(e) => onShowQuickHelp?.(e, '仮置き場', 'コマを一時退避する場所です。ログイン中のGoogleアカウント専用の仮置き場です。')}
          onMouseLeave={() => onHideQuickHelp?.()}
        >
          <div className="flex items-center gap-2 text-xs font-bold text-slate-600">
            <div className="p-1 bg-indigo-100 text-indigo-600 rounded">
              <ClipboardList size={14} />
            </div>
            <span>仮置き場（あなた専用）</span>
            <span className="text-[10px] px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-600 border border-indigo-100">
              {activeTempItems.length}件
            </span>
          </div>
          <span className="text-[10px] text-slate-400 font-medium">Googleアカウント別</span>
        </div>

        <div className="flex-1 overflow-y-auto p-3">
          {activeTempItems.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-400 border-2 border-dashed border-slate-200 rounded-xl">
              <p className="text-xs font-medium">ここにドロップ</p>
            </div>
          ) : (
            <div className="grid grid-cols-2 gap-3">
              {activeTempItems.map((item) => {
                const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
                const hoverCodeText = (item.code || '').trim();
                return (
                  <div
                    key={item.id}
                    className="group relative border border-slate-200 rounded-lg p-2 bg-white hover:shadow-md cursor-grab active:cursor-grabbing flex items-center justify-center min-h-[80px] transition-all"
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
                    <div className="relative w-full h-16 overflow-hidden rounded">
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
                      <div className="absolute top-1 left-1 bg-slate-800/80 backdrop-blur-sm text-white text-[9px] px-1.5 py-0.5 rounded shadow-sm">
                        {item.label}
                      </div>
                    )}
                    {hoverCodeText && (
                      <div className="absolute left-1 right-1 bottom-1 pointer-events-none opacity-0 group-hover:opacity-100 transition-opacity duration-150">
                        <div className="mx-auto max-w-full truncate text-[10px] font-mono font-bold text-white bg-slate-900/85 px-2 py-1 rounded shadow-md text-center">
                          {hoverCodeText}
                        </div>
                      </div>
                    )}
                    <button
                      onClick={() => onDeleteFromTemp(item.id)}
                      className="absolute -top-2 -right-2 bg-white rounded-full p-1 shadow-md text-rose-500 hover:bg-rose-50 opacity-0 group-hover:opacity-100 transition-all border border-slate-100"
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

      {previewImage && (
        <div
          className="fixed inset-0 z-[170] bg-black/70 flex items-center justify-center p-5"
          onClick={() => setPreviewImage(null)}
          title="クリックでプレビューを閉じる"
        >
          <div className="max-w-[92vw] max-h-[90vh] rounded-2xl border border-white/20 bg-slate-950/90 p-3 shadow-2xl">
            <img
              src={previewImage.src}
              alt={previewImage.name || 'preview'}
              className="max-w-[86vw] max-h-[80vh] object-contain rounded-lg"
            />
            {previewImage.name && (
              <p className="mt-2 text-[11px] text-slate-200 font-mono truncate text-center">
                {previewImage.name}
              </p>
            )}
          </div>
        </div>
      )}

      <div
        className="absolute right-0 top-0 bottom-0 w-1 cursor-ew-resize hover:bg-indigo-400 transition-colors z-30"
        onMouseDown={() => setResizing(true)}
      />
    </div>
  );
});


export default Sidebar;
