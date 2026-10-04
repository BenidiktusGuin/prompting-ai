import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx';

const MAX_CHARS_PER_SHEET = 2000;

export default function App() {
  // State Dokumen
  const [title, setTitle] = useState(() => {
    return localStorage.getItem('manuscript_title') || '';
  });

  const [sheets, setSheets] = useState(() => {
    try {
      const stored = localStorage.getItem('manuscript_sheets');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
      // Cek legacy draft
      const oldDraft = localStorage.getItem('manuscript_draft');
      if (oldDraft) {
        if (oldDraft.length > MAX_CHARS_PER_SHEET) {
          const chunks = [];
          for (let i = 0; i < oldDraft.length; i += MAX_CHARS_PER_SHEET) {
            chunks.push(oldDraft.slice(i, i + MAX_CHARS_PER_SHEET));
          }
          return chunks;
        }
        return [oldDraft];
      }
    } catch (e) {
      console.warn('Gagal memuat naskah dari storage:', e);
    }
    return [''];
  });

  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [animatingSheetIndex, setAnimatingSheetIndex] = useState(null);
  const [toast, setToast] = useState({ message: '', visible: false });
  const [isSaving, setIsSaving] = useState(false);

  // Refs untuk navigasi kursor & fokus
  const activeTextareaRef = useRef(null);
  const titleInputRef = useRef(null);
  const toastTimeoutRef = useRef(null);
  const pendingFocusRef = useRef(null); // { index, position: 'start' | 'end' | number }

  // Simpan otomatis ke localStorage
  useEffect(() => {
    localStorage.setItem('manuscript_title', title);
    localStorage.setItem('manuscript_sheets', JSON.stringify(sheets));
    localStorage.setItem('manuscript_draft', sheets.join(''));
  }, [title, sheets]);

  // Efek fokus saat activeSheetIndex berubah
  useEffect(() => {
    if (pendingFocusRef.current) {
      const { position } = pendingFocusRef.current;
      pendingFocusRef.current = null;
      if (activeTextareaRef.current) {
        activeTextareaRef.current.focus();
        const len = activeTextareaRef.current.value.length;
        if (position === 'start') {
          activeTextareaRef.current.setSelectionRange(0, 0);
        } else if (position === 'end') {
          activeTextareaRef.current.setSelectionRange(len, len);
        } else if (typeof position === 'number') {
          activeTextareaRef.current.setSelectionRange(position, position);
        }
      }
    }
  }, [activeSheetIndex]);

  // Tampilkan notifikasi toast
  const showToast = useCallback((msg) => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToast({ message: msg, visible: true });
    toastTimeoutRef.current = setTimeout(() => {
      setToast(prev => ({ ...prev, visible: false }));
    }, 3200);
  }, []);

  // Hitung total karakter & kata seluruh dokumen (di luar judul)
  const totalChars = sheets.reduce((acc, curr) => acc + curr.length, 0);
  const totalWords = sheets.reduce((acc, curr) => {
    const words = curr.trim().split(/\s+/).filter(w => w.length > 0);
    return acc + words.length;
  }, 0);

  // Seimbangkan karakter maju ke lembar-lembar berikutnya jika melebihi 2000
  const balanceSheets = (sheetsArray, startIndex) => {
    const result = [...sheetsArray];
    let idx = startIndex;
    while (idx < result.length && result[idx].length > MAX_CHARS_PER_SHEET) {
      const text = result[idx];
      result[idx] = text.slice(0, MAX_CHARS_PER_SHEET);
      const overflow = text.slice(MAX_CHARS_PER_SHEET);

      if (idx + 1 < result.length) {
        result[idx + 1] = overflow + result[idx + 1];
      } else {
        result.push(overflow);
      }
      idx++;
    }
    return result;
  };

  // Penanganan Pengetikan: Batasi 2.000 karakter per lembar
  const handleSheetInput = (e, index) => {
    const val = e.target.value;

    if (val.length > MAX_CHARS_PER_SHEET) {
      // Potong tepat 2000 karakter untuk lembar saat ini
      const currentSheetText = val.slice(0, MAX_CHARS_PER_SHEET);
      const overflowText = val.slice(MAX_CHARS_PER_SHEET);

      let nextSheets = [...sheets];
      nextSheets[index] = currentSheetText;

      const nextIndex = index + 1;
      if (nextIndex < nextSheets.length) {
        nextSheets[nextIndex] = overflowText + nextSheets[nextIndex];
      } else {
        nextSheets.push(overflowText);
      }

      nextSheets = balanceSheets(nextSheets, nextIndex);

      setSheets(nextSheets);
      setActiveSheetIndex(nextIndex);
      setAnimatingSheetIndex(nextIndex);
      pendingFocusRef.current = { index: nextIndex, position: overflowText.length };

      showToast(`Mencapai 2.000 karakter. Lembar baru dibuka!`);
      return;
    }

    const nextSheets = [...sheets];
    nextSheets[index] = val;
    setSheets(nextSheets);
  };

  // Penanganan Paste teks panjang
  const handleSheetPaste = (e, index) => {
    const clipboardData = e.clipboardData;
    if (!clipboardData) return;

    const pastedText = clipboardData.getData('text');
    const textarea = e.target;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const currentVal = textarea.value;

    const projected = currentVal.slice(0, start) + pastedText + currentVal.slice(end);

    if (projected.length > MAX_CHARS_PER_SHEET) {
      e.preventDefault();

      let nextSheets = [...sheets];
      nextSheets[index] = projected.slice(0, MAX_CHARS_PER_SHEET);
      let remaining = projected.slice(MAX_CHARS_PER_SHEET);

      let targetIndex = index;
      while (remaining.length > 0) {
        targetIndex++;
        const chunk = remaining.slice(0, MAX_CHARS_PER_SHEET);
        remaining = remaining.slice(MAX_CHARS_PER_SHEET);

        if (targetIndex < nextSheets.length) {
          nextSheets[targetIndex] = chunk + nextSheets[targetIndex];
        } else {
          nextSheets.push(chunk);
        }
      }

      nextSheets = balanceSheets(nextSheets, index + 1);

      setSheets(nextSheets);
      setActiveSheetIndex(targetIndex);
      setAnimatingSheetIndex(targetIndex);
      pendingFocusRef.current = { index: targetIndex, position: 'end' };

      showToast(`Teks didistribusikan ke ${nextSheets.length} lembar kerja`);
    }
  };

  // Penanganan Tombol Backspace di awal lembar
  const handleSheetKeyDown = (e, index) => {
    if (e.key === 'Backspace' && e.target.selectionStart === 0 && e.target.selectionEnd === 0) {
      if (index > 0) {
        e.preventDefault();
        // Jika lembar saat ini kosong dan lembar terakhir, hapus otomatis
        if (sheets[index].length === 0 && index === sheets.length - 1) {
          const nextSheets = sheets.slice(0, -1);
          setSheets(nextSheets);
          const prevIdx = index - 1;
          setActiveSheetIndex(prevIdx);
          pendingFocusRef.current = { index: prevIdx, position: 'end' };
          showToast(`Kembali ke Lembar ${prevIdx + 1}`);
        } else {
          // Pindah ke akhir lembar sebelumnya
          setActiveSheetIndex(index - 1);
          pendingFocusRef.current = { index: index - 1, position: 'end' };
        }
      }
    }
  };

  // Navigasi Lembar
  const goToSheet = (newIndex, focusPos = 'end') => {
    if (newIndex < 0 || newIndex >= sheets.length) return;
    if (newIndex > activeSheetIndex) {
      setAnimatingSheetIndex(newIndex);
    }
    setActiveSheetIndex(newIndex);
    pendingFocusRef.current = { index: newIndex, position: focusPos };
  };

  // Tambah Lembar Manual
  const handleAddSheet = () => {
    const nextSheets = [...sheets, ''];
    const newIdx = nextSheets.length - 1;
    setSheets(nextSheets);
    setActiveSheetIndex(newIdx);
    setAnimatingSheetIndex(newIdx);
    pendingFocusRef.current = { index: newIdx, position: 'start' };
    showToast(`Lembar ${newIdx + 1} ditambahkan`);
  };

  // Hapus Lembar Kosong
  const handleRemoveSheet = (index) => {
    if (sheets.length <= 1) return;
    const nextSheets = sheets.filter((_, i) => i !== index);
    const newActive = Math.min(activeSheetIndex, nextSheets.length - 1);
    setSheets(nextSheets);
    setActiveSheetIndex(newActive);
    pendingFocusRef.current = { index: newActive, position: 'end' };
    showToast(`Lembar ${index + 1} dihapus`);
  };

  // Simpan File (.docx)
  const saveFile = useCallback(async () => {
    if (isSaving) return;
    setIsSaving(true);

    try {
      const cleanTitle = title.trim();
      const docChildren = [];

      // Judul Naskah di bagian atas jika diisi
      if (cleanTitle) {
        docChildren.push(
          new Paragraph({
            text: cleanTitle,
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 360 },
          })
        );
      }

      // Setiap lembar kerja dimasukkan sebagai paragraf di Word
      sheets.forEach((sheetText, sheetIndex) => {
        const lines = sheetText.split('\n');

        lines.forEach((line, lineIndex) => {
          // Berikan page break sebelum lembar baru (mulai Lembar 2)
          const needsPageBreak = sheetIndex > 0 && lineIndex === 0;

          docChildren.push(
            new Paragraph({
              pageBreakBefore: needsPageBreak,
              alignment: AlignmentType.LEFT,
              spacing: {
                after: 140,
                line: 360, // Spasi baris 1.5
              },
              children: [
                new TextRun({
                  text: line,
                  font: 'Calibri',
                  size: 24, // 12pt
                  color: '1f2937',
                }),
              ],
            })
          );
        });
      });

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 1440, // 1 inch
                  right: 1440,
                  bottom: 1440,
                  left: 1440,
                },
              },
            },
            children: docChildren.length > 0 ? docChildren : [new Paragraph('')],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);

      let filename = 'naskah.docx';
      const baseName = cleanTitle || sheets[0].trim().split('\n')[0].trim();
      if (baseName.length > 0) {
        const sanitized = baseName.replace(/[/\\?%*:|"<>]/g, '').trim().substring(0, 35);
        if (sanitized) filename = `${sanitized}.docx`;
      }

      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      showToast(`Tersimpan (${sheets.length} lembar) sebagai "${filename}"`);
    } catch (err) {
      console.error('Gagal membuat file DOCX:', err);
      showToast('Gagal membuat file .docx. Silakan coba lagi.');
    } finally {
      setIsSaving(false);
    }
  }, [title, sheets, isSaving, showToast]);

  // Pintasan Keyboard Global
  useEffect(() => {
    const handleGlobalKeyDown = (e) => {
      // Simpan Ctrl+S
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        saveFile();
      }

      // Lembar Sebelumnya: Alt+Left atau Ctrl+PageUp
      if ((e.altKey && e.key === 'ArrowLeft') || (e.ctrlKey && e.key === 'PageUp')) {
        e.preventDefault();
        if (activeSheetIndex > 0) {
          goToSheet(activeSheetIndex - 1);
        }
      }

      // Lembar Berikutnya: Alt+Right atau Ctrl+PageDown
      if ((e.altKey && e.key === 'ArrowRight') || (e.ctrlKey && e.key === 'PageDown')) {
        e.preventDefault();
        if (activeSheetIndex < sheets.length - 1) {
          goToSheet(activeSheetIndex + 1);
        }
      }
    };

    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, [activeSheetIndex, sheets.length, saveFile]);

  return (
    <>
      {/* Top Sticky Navigation Bar */}
      <header className="top-action-bar">
        {/* Statistik Keseluruhan Dokumen */}
        <div className="doc-stats-pill" title="Total karakter & kata seluruh lembar (diluar judul)">
          <span>Total:</span>
          <span className="highlight">{totalChars.toLocaleString('id-ID')}</span> huruf
          <span>·</span>
          <span className="highlight">{totalWords.toLocaleString('id-ID')}</span> kata
        </div>

        {/* Kontrol Navigasi Antar Lembar */}
        <div className="sheet-nav-controls">
          <button
            type="button"
            className="btn-nav"
            onClick={() => goToSheet(activeSheetIndex - 1)}
            disabled={activeSheetIndex === 0}
            title="Lembar Sebelumnya (Alt + ←)"
            aria-label="Lembar Sebelumnya"
          >
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5 8.25 12l7.5-7.5" />
            </svg>
          </button>

          <div className="sheet-badge-indicator">
            Lembar {activeSheetIndex + 1} / {sheets.length}
          </div>

          <button
            type="button"
            className="btn-nav"
            onClick={() => goToSheet(activeSheetIndex + 1)}
            disabled={activeSheetIndex >= sheets.length - 1}
            title="Lembar Berikutnya (Alt + →)"
            aria-label="Lembar Berikutnya"
          >
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" d="m8.25 4.5 7.5 7.5-7.5 7.5" />
            </svg>
          </button>

          <button
            type="button"
            className="btn-add-sheet"
            onClick={handleAddSheet}
            title="Tambah lembar kerja baru secara manual"
          >
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
            </svg>
            <span>Lembar Baru</span>
          </button>
        </div>

        {/* Tombol Simpan Naskah (.docx) */}
        <button
          type="button"
          className="btn-save"
          onClick={saveFile}
          disabled={isSaving}
          title="Simpan naskah ke format Microsoft Word (.docx) (Ctrl+S)"
        >
          {isSaving ? (
            <>
              <svg className="animate-spin" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0 3.181 3.183a8.25 8.25 0 0 0 13.803-3.7M4.031 9.865a8.25 8.25 0 0 1 13.803-3.7l3.181 3.182m0-4.991v4.99" />
              </svg>
              <span>Menyimpan...</span>
            </>
          ) : (
            <>
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round"
                  d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m.75 12 3 3m0 0 3-3m-3 3v-6m-1.5-9H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z" />
              </svg>
              <span>Save File (.docx)</span>
            </>
          )}
        </button>
      </header>

      {/* 3D Workspace Stage */}
      <main className="workspace-stage">
        <div className="sheets-stack">
          {sheets.map((sheetText, index) => {
            const isActive = index === activeSheetIndex;
            const isBehind = index < activeSheetIndex;
            const isAhead = index > activeSheetIndex;
            const isAnimating = animatingSheetIndex === index;

            // Menentukan kelas kedalaman dan efek blur
            let cardClasses = 'sheet-card';
            if (isActive) {
              cardClasses += ' is-active';
              if (isAnimating) cardClasses += ' animate-slide-in';
            } else if (isBehind) {
              cardClasses += ' is-behind';
              const depth = activeSheetIndex - index;
              if (depth === 1) cardClasses += ' is-depth-1';
              else if (depth === 2) cardClasses += ' is-depth-2';
              else cardClasses += ' is-depth-deep';
            } else if (isAhead) {
              cardClasses += ' is-ahead';
            }

            const charCount = sheetText.length;
            const pct = Math.min((charCount / MAX_CHARS_PER_SHEET) * 100, 100);

            let charCounterClass = 'sheet-char-counter';
            if (charCount >= MAX_CHARS_PER_SHEET) charCounterClass += ' is-max-limit';
            else if (charCount >= MAX_CHARS_PER_SHEET * 0.9) charCounterClass += ' is-near-limit';

            let progressFillClass = 'progress-bar-fill';
            if (pct >= 100) progressFillClass += ' fill-limit';
            else if (pct >= 90) progressFillClass += ' fill-warning';

            return (
              <article
                key={index}
                className={cardClasses}
                onClick={isBehind ? () => goToSheet(index) : undefined}
                onAnimationEnd={() => {
                  if (animatingSheetIndex === index) setAnimatingSheetIndex(null);
                }}
              >
                {/* Hint saat mengarahkan kursor ke lembar belakang yang buram */}
                {isBehind && (
                  <div className="sheet-behind-hint">
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <path d="M15 18l-6-6 6-6" />
                    </svg>
                    <span>Buka Lembar {index + 1}</span>
                  </div>
                )}

                {/* 1. Header Internal Lembar */}
                <div className="sheet-internal-header">
                  <div className="sheet-number-tag">
                    <span className="dot"></span>
                    <span>Lembar {index + 1}</span>
                  </div>

                  <div className="sheet-char-progress-wrapper">
                    <span className={charCounterClass}>
                      {charCount.toLocaleString('id-ID')} / {MAX_CHARS_PER_SHEET.toLocaleString('id-ID')} huruf
                    </span>
                    <div className="progress-bar-bg">
                      <div className={progressFillClass} style={{ width: `${pct}%` }}></div>
                    </div>
                  </div>
                </div>

                {/* 2. Judul Dokumen (Lembar 1) vs Banner Lanjutan (Lembar 2+) */}
                {index === 0 ? (
                  <div className="editor-title-container">
                    <input
                      ref={titleInputRef}
                      type="text"
                      className="editor-title"
                      placeholder="Judul Naskah..."
                      autoComplete="off"
                      spellCheck={false}
                      value={title}
                      onChange={(e) => setTitle(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          if (activeTextareaRef.current) activeTextareaRef.current.focus();
                        }
                      }}
                    />
                  </div>
                ) : (
                  <div className="continuation-banner">
                    <div className="continuation-info">
                      <span className="continuation-pill">Lanjutan</span>
                      <span className="continuation-title-preview">
                        {title.trim() ? title : 'Naskah Tanpa Judul'}
                      </span>
                    </div>
                    <button
                      type="button"
                      className="btn-edit-title-link"
                      onClick={(e) => {
                        e.stopPropagation();
                        goToSheet(0);
                        setTimeout(() => {
                          if (titleInputRef.current) {
                            titleInputRef.current.focus();
                            titleInputRef.current.select();
                          }
                        }, 120);
                      }}
                    >
                      Edit Judul di Lembar 1 ↗
                    </button>
                  </div>
                )}

                {/* 3. Textarea Isi Lembar */}
                <textarea
                  ref={isActive ? activeTextareaRef : null}
                  className="editor-textarea"
                  placeholder={
                    index === 0
                      ? 'Mulai ketik isi naskah Anda di sini...'
                      : `Lanjutan naskah untuk Lembar ${index + 1}...`
                  }
                  spellCheck={false}
                  value={sheetText}
                  readOnly={!isActive}
                  tabIndex={isActive ? 0 : -1}
                  onChange={(e) => handleSheetInput(e, index)}
                  onKeyDown={(e) => handleSheetKeyDown(e, index)}
                  onPaste={(e) => handleSheetPaste(e, index)}
                />

                {/* 4. Opsi Hapus Lembar Kosong Terakhir */}
                {index > 0 && index === sheets.length - 1 && sheetText.trim() === '' && (
                  <div className="sheet-card-footer">
                    <button
                      type="button"
                      className="btn-delete-sheet"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveSheet(index);
                      }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" d="m14.74 9-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 0 1-2.244 2.077H8.084a2.25 2.25 0 0 1-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 0 0-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 0 1 3.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 0 0-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 0 0-7.5 0" />
                      </svg>
                      <span>Hapus Lembar Kosong Ini</span>
                    </button>
                  </div>
                )}
              </article>
            );
          })}
        </div>
      </main>

      {/* Notifikasi Toast */}
      <div
        className={`toast ${toast.visible ? 'show' : ''}`}
        role="status"
        aria-live="polite"
      >
        {toast.message}
      </div>
    </>
  );
}
