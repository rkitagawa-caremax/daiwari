import React, { useCallback, useState } from 'react';
import { ClipboardList, X } from 'lucide-react';

import ImagePreviewModal from '../../components/dialogs/ImagePreviewModal';
import {
  clearActiveNativeDragPayload,
  getDragPayload,
  isDropEventHandled,
  markDropEventHandled,
  setDragPayload
} from '../../lib/dragPayload';
import { FreeLabelPreview } from './Sidebar';

// コントロールパネル下に置くコンパクト版の仮置き場。
// ドロップ受付は data-daiwari-dropzone-id="temp" (ポインタDnD) と native DnD の両対応。
const TempShelfPanel = React.memo(({
  tempItems,
  imageDataById,
  onDeleteFromTemp,
  onApplyDragPayloadToTemp,
  onStartPointerDrag,
  onShowQuickHelp,
  onHideQuickHelp
}) => {
  const [previewImage, setPreviewImage] = useState(null);
  const items = tempItems || [];

  const handleTogglePreview = useCallback((src, name = '') => {
    if (!src) return;
    setPreviewImage((prev) => (prev?.src === src ? null : { src, name }));
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

  return (
    <div
      className="flex w-full flex-col overflow-hidden rounded-2xl border border-slate-200/80 bg-white/90 shadow-lg shadow-slate-300/25 backdrop-blur-md"
      data-daiwari-dropzone-id="temp"
      onDragOverCapture={(e) => e.preventDefault()}
      onDropCapture={handleDropToTemp}
      onDragOver={(e) => e.preventDefault()}
      onDrop={handleDropToTemp}
    >
      <div
        className="flex items-center justify-between border-b border-slate-200/70 px-2.5 py-1.5"
        onMouseEnter={(e) => onShowQuickHelp?.(e, '仮置き場', 'コマを一時退避する場所です。ログイン中のGoogleアカウント専用の仮置き場です。')}
        onMouseLeave={() => onHideQuickHelp?.()}
      >
        <div className="flex min-w-0 items-center gap-1.5 text-[11px] font-bold text-slate-700">
          <ClipboardList size={13} className="flex-shrink-0 text-indigo-500" />
          <span className="truncate">仮置き場</span>
          <span className="rounded-full bg-indigo-50 px-1.5 py-0.5 text-[9px] text-indigo-600">
            {items.length}
          </span>
        </div>
        <span className="text-[9px] font-medium text-slate-400">自分専用</span>
      </div>

      <div className="max-h-96 min-h-40 flex-1 overflow-y-auto p-1.5">
        {items.length === 0 ? (
          <div className="flex h-36 flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 text-slate-400">
            <p className="text-[10px] font-medium">ここにドロップ</p>
          </div>
        ) : (
          <div className="grid grid-cols-2 gap-1.5">
            {items.map((item) => {
              const resolvedImg = item.image || (item.imageId ? imageDataById?.[item.imageId] : null);
              const hoverCodeText = (item.code || '').trim();
              return (
                <div
                  key={item.id}
                  className="group relative flex cursor-grab flex-col items-center rounded-lg border border-slate-200 bg-white p-1 transition-all hover:border-slate-300 hover:shadow-sm active:cursor-grabbing"
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
                  <div className="relative h-10 w-full overflow-hidden rounded-md bg-slate-50">
                    {resolvedImg ? (
                      <img src={resolvedImg} alt="temp" className="h-full w-full object-contain" draggable={false} />
                    ) : (
                      <div className="flex h-full w-full flex-col items-center justify-center bg-slate-50 text-slate-400">
                        <span className="font-mono text-[9px]">{item.code || 'No Image'}</span>
                      </div>
                    )}
                    <FreeLabelPreview item={item} />
                  </div>

                  {item.label && (
                    <div className="absolute left-0.5 top-0.5 max-w-[70%] truncate rounded bg-slate-800/75 px-1 py-0.5 text-[7px] text-white">
                      {item.label}
                    </div>
                  )}
                  <p className="mt-0.5 w-full truncate text-center font-mono text-[8px] font-bold text-slate-600">
                    {hoverCodeText || item.originalName || '仮置き'}
                  </p>
                  <button
                    onClick={(event) => {
                      event.stopPropagation();
                      onDeleteFromTemp(item.id);
                    }}
                    onPointerDown={(event) => event.stopPropagation()}
                    className="absolute right-0.5 top-0.5 rounded-full border border-slate-200 bg-white/90 p-0.5 text-slate-400 shadow-sm transition-colors hover:border-rose-200 hover:bg-rose-50 hover:text-rose-600"
                    title="仮置き場から削除"
                  >
                    <X size={10} strokeWidth={3} />
                  </button>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <ImagePreviewModal preview={previewImage} onClose={() => setPreviewImage(null)} />
    </div>
  );
});

TempShelfPanel.displayName = 'TempShelfPanel';

export default TempShelfPanel;
