import { useCallback, useEffect, useMemo, useState, type RefObject } from 'react';
import { X } from 'lucide-react';
import {
  FORMAT_CATEGORIES,
  FormatCategoryId,
  allFormatPresets,
  applyCellFormatToWorkbook,
  categoryForFormat,
  deleteCustomFormat,
  formatPreview,
  isValidFormatCode,
  loadCustomFormats,
  saveCustomFormat,
  sanitizeFormatCode,
} from '../sheet/cellFormatUtils';

type Props = {
  open: boolean;
  workbookRef: RefObject<any>;
  onClose: () => void;
  onApplied?: () => void;
};

export default function FormatCellsDialog({ open, workbookRef, onClose, onApplied }: Props) {
  const [categoryId, setCategoryId] = useState<FormatCategoryId>('general');
  const [typeCode, setTypeCode] = useState('General');
  const [sampleRaw, setSampleRaw] = useState<unknown>(null);
  const [customListTick, setCustomListTick] = useState(0);
  const [applyError, setApplyError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setApplyError(null);
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [open, onClose]);

  useEffect(() => {
    const workbook = workbookRef.current;
    if (!open || !workbook) return;
    try {
      const sel = workbook.getSelection?.()?.[0];
      if (!sel) return;
      const r = sel.row[0];
      const c = sel.column[0];
      const sheet = workbook.getSheet?.();
      const cell = sheet?.data?.[r]?.[c];
      const fa = cell?.ct?.fa || 'General';
      setTypeCode(fa);
      setCategoryId(categoryForFormat(fa));
      setSampleRaw(cell?.v ?? workbook.getCellValue?.(r, c, { type: 'v' }));
    } catch {
      setTypeCode('General');
      setCategoryId('general');
      setSampleRaw(null);
    }
  }, [open, workbookRef]);

  const presets = useMemo(() => {
    void customListTick;
    if (categoryId === 'custom') return allFormatPresets();
    const cat = FORMAT_CATEGORIES.find((c) => c.id === categoryId);
    return cat?.formats ?? [];
  }, [categoryId, customListTick]);

  const sampleText = useMemo(() => formatPreview(typeCode, sampleRaw), [typeCode, sampleRaw]);

  const selectFormat = useCallback((fa: string) => {
    setTypeCode(fa);
    setCategoryId(categoryForFormat(fa));
  }, []);

  const handleOk = useCallback(() => {
    const workbook = workbookRef.current;
    const fa = typeCode.trim() || 'General';
    if (!workbook?.getSelection?.()?.length) {
      setApplyError('Сначала выделите одну или несколько ячеек на листе.');
      return;
    }
    const safeFa = sanitizeFormatCode(fa);
    if (!isValidFormatCode(safeFa)) {
      setApplyError('Некорректный код формата. Проверьте синтаксис (символ ₽ нужно брать в кавычки: #" ₽").');
      return;
    }
    if (applyCellFormatToWorkbook(workbook, safeFa)) {
      saveCustomFormat(safeFa);
      onApplied?.();
      onClose();
      return;
    }
    setApplyError('Не удалось применить формат. Проверьте код формата.');
  }, [workbookRef, typeCode, onApplied, onClose]);

  const handleDelete = useCallback(() => {
    if (!loadCustomFormats().includes(typeCode)) return;
    deleteCustomFormat(typeCode);
    setCustomListTick((n) => n + 1);
  }, [typeCode]);

  if (!open) return null;

  const canDelete = loadCustomFormats().includes(typeCode);

  return (
    <div
      className="tw-format-cells-backdrop"
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div className="tw-format-cells-dialog" role="dialog" aria-modal="true" aria-label="Формат ячеек">
        <div className="tw-format-cells-header">
          <h3>Формат ячеек</h3>
          <button type="button" className="tw-format-cells-close" onClick={onClose} aria-label="Закрыть">
            <X size={18} />
          </button>
        </div>

        <div className="tw-format-cells-tabs">
          <span className="tw-format-cells-tab active">Число</span>
        </div>

        <div className="tw-format-cells-body">
          <div className="tw-format-cells-categories">
            <div className="tw-format-cells-label">Числовые форматы:</div>
            <ul className="tw-format-cells-cat-list">
              {FORMAT_CATEGORIES.map((cat) => (
                <li key={cat.id}>
                  <button
                    type="button"
                    className={categoryId === cat.id ? 'active' : ''}
                    onClick={() => {
                      setCategoryId(cat.id);
                      if (cat.formats[0]) setTypeCode(cat.formats[0].fa);
                    }}
                  >
                    {cat.label}
                  </button>
                </li>
              ))}
            </ul>
          </div>

          <div className="tw-format-cells-main">
            <div className="tw-format-cells-sample-row">
              <span className="tw-format-cells-label">Образец:</span>
              <div className="tw-format-cells-sample">{sampleText || ' '}</div>
            </div>

            <div className="tw-format-cells-type-row">
              <label className="tw-format-cells-label" htmlFor="tw-format-type-input">Тип:</label>
              <input
                id="tw-format-type-input"
                className="tw-format-cells-type-input"
                value={typeCode}
                onChange={(e) => {
                  setTypeCode(e.target.value);
                  setCategoryId('custom');
                }}
                spellCheck={false}
              />
            </div>

            <div className="tw-format-cells-preset-list-wrap">
              <ul className="tw-format-cells-preset-list">
                {presets.map((p) => (
                  <li key={p.fa}>
                    <button
                      type="button"
                      className={typeCode === p.fa ? 'active' : ''}
                      onClick={() => selectFormat(p.fa)}
                    >
                      {p.label}
                    </button>
                  </li>
                ))}
              </ul>
              <button
                type="button"
                className="tw-format-cells-delete"
                disabled={!canDelete}
                onClick={handleDelete}
              >
                Удалить
              </button>
            </div>

            <p className="tw-format-cells-hint">
              Введите код числового формата, используя один из существующих кодов в качестве образца.
              Поддерживаются коды Excel и русские обозначения (например, ДД.ММ.ГГГГ ч:мм).
            </p>
            {applyError && <p className="tw-format-cells-error">{applyError}</p>}
          </div>
        </div>

        <div className="tw-format-cells-footer">
          <button type="button" className="tw-format-cells-btn primary" onClick={handleOk}>OK</button>
          <button type="button" className="tw-format-cells-btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
