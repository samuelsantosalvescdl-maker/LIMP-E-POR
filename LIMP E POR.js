/*********************************************************
 *  ALMOX — FRACIONAMENTO + ETIQUETAS + ANTI MÃO-BOBA
 *
 *  ✅ SEM "CONFIRMAÇÃO DE COMANDOS" / SEM rodar como proprietário
 *  ✅ Sequencial LIF em DADOSFRACIONAMENTOS (col J -> escreve em P)
 *  ✅ Impressão (ETIQUETAS) + PDF + limpeza + nota em H
 *  ✅ Anti Mão-Boba: bloqueia o proprietário nas áreas protegidas
 *  ✅ Scripts SEMPRE podem escrever (sem virar "inválido" depois)
 *  ✅ Trigger diário 03:00 para reforçar bloqueio e atualizar snapshots
 *********************************************************/

/* =====================================================================
 * 0) CONFIG ANTI
 * ===================================================================== */
const ALMOX_ANTI = {
  CONFIG_SHEET: '__ANTI_MAO_BOBA__',
  EDIT_MODE_A1: 'A1',
  EDIT_MODE_NAMED: 'ANTI_EDIT_MODE',

  SNAP_PREFIX: '__ANTI_SNAP__',
  HELP_TEXT: 'Bloqueado (Anti Mão-Boba). Use o menu para liberar por 10 min.',

  UNLOCK_MINUTES: 10,

  // Segurança (evita snapshot “gigante” em abas que foram esticadas por acidente)
  SAFE_MAX_ROW: 8000,
  SAFE_MAX_COL: 300,

  // Chunk para copiar valores pro snapshot (reduz Service error)
  ROW_CHUNK: 120,
  COL_CHUNK: 35,

  RETRIES: 10
};

/* =====================================================================
 * 0.1) WRAPPER — Scripts sempre podem escrever SEM ficar inválido depois
 *  - Uso:
 *      almox_anti_runUnlocked_(ss, () => {
 *         // ... escreve em ranges ...
 *      }, [range1, range2]);
 *  - Ele:
 *      1) liga ANTI_EDIT_MODE (aceita qualquer valor)
 *      2) executa a escrita
 *      3) atualiza SOMENTE o snapshot dos ranges alterados (leve)
 *      4) volta ANTI_EDIT_MODE ao estado anterior
 * ===================================================================== */
function almox_anti_runUnlocked_(ss, fn, touchRanges) {
  ss = ss || SpreadsheetApp.getActive();
  const modeCell = ss.getRangeByName(ALMOX_ANTI.EDIT_MODE_NAMED);
  if (!modeCell) throw new Error('ANTI_EDIT_MODE não encontrado. Rode almox_anti_setup().');

  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  const prev = !!modeCell.getValue();
  try {
    modeCell.setValue(true);
    const ret = fn ? fn() : null;

    // Atualiza snapshot apenas dos ranges alterados
    if (touchRanges && touchRanges.length) {
      almox_anti_updateSnapshotForRanges_(ss, touchRanges);
    }

    return ret;
  } finally {
    try { modeCell.setValue(prev); } catch (_) {}
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Atualiza snapshot só para ranges específicos (bem mais leve que refresh geral). */
function almox_anti_updateSnapshotForRanges_(ss, ranges) {
  ss = ss || SpreadsheetApp.getActive();
  const bySheet = new Map();

  (ranges || []).forEach(rg => {
    if (!rg) return;
    const sh = rg.getSheet();
    const id = sh.getSheetId();
    if (!bySheet.has(id)) bySheet.set(id, { sheet: sh, ranges: [] });
    bySheet.get(id).ranges.push(rg);
  });

  for (const entry of bySheet.values()) {
    const sh = entry.sheet;
    const snap = almox_anti_getSnapSheet_(ss, sh, { createIfMissing: true });

    // Garante tamanho do snapshot até o máximo necessário
    let maxR = 1, maxC = 1;
    entry.ranges.forEach(rg => {
      maxR = Math.max(maxR, rg.getLastRow());
      maxC = Math.max(maxC, rg.getLastColumn());
    });
    almox_anti_ensureSize_(snap, maxR, maxC);

    // Copia valores em chunks (reduz Service error)
    entry.ranges.forEach(rg => {
      almox_anti_copyValuesChunked_(rg, snap.getRange(rg.getRow(), rg.getColumn(), rg.getNumRows(), rg.getNumColumns()));
    });
  }
}

function almox_anti_copyValuesChunked_(srcRange, dstRangeSameSize) {
  const rows = srcRange.getNumRows();
  const cols = srcRange.getNumColumns();
  const sr = srcRange.getRow();
  const sc = srcRange.getColumn();

  for (let rOff = 0; rOff < rows; rOff += ALMOX_ANTI.ROW_CHUNK) {
    const h = Math.min(ALMOX_ANTI.ROW_CHUNK, rows - rOff);

    for (let cOff = 0; cOff < cols; cOff += ALMOX_ANTI.COL_CHUNK) {
      const w = Math.min(ALMOX_ANTI.COL_CHUNK, cols - cOff);

      const src = srcRange.getSheet().getRange(sr + rOff, sc + cOff, h, w);
      const dst = dstRangeSameSize.getSheet().getRange(sr + rOff, sc + cOff, h, w);

      almox_anti_retry_(() => {
        src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      }, 'snapshot.chunk.copyTo');
    }
    Utilities.sleep(10);
  }
}

/* =====================================================================
 * 0.2) Trigger diário 03:00
 * ===================================================================== */
function almox_anti_installDailyTrigger_() {
  const FN = 'almox_anti_rotinaDiaria_';

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && t.getHandlerFunction() === FN)
    .forEach(t => { try { ScriptApp.deleteTrigger(t); } catch (_) {} });

  ScriptApp.newTrigger(FN)
    .timeBased()
    .atHour(3)       // 03:00
    .everyDays(1)
    .create();
}

/** Rotina diária (silenciosa): trava, atualiza snapshots e reaplica bloqueio */
function almox_anti_rotinaDiaria_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const cell = ss.getRangeByName(ALMOX_ANTI.EDIT_MODE_NAMED);
    if (cell) cell.setValue(false);

    almox_anti_refreshAllSnapshots_();
    almox_anti_aplicarBloqueio(true); // silent
  } catch (e) {
    console.error('almox_anti_rotinaDiaria_:', e);
  }
}

/* =====================================================================
 * 1) SEQUENCIAL LIF — DADOSFRACIONAMENTOS
 * ===================================================================== */
const CFG_LIF = {
  NOME_INTERVALO: 'DADOSFRACIONAMENTOS',
  COL_EDIT: 10,       // J
  COL_SAIDA: 16,      // P
  PREFIXO: 'LIF ',
  DIGITOS: 8,
  PROP_KEY: 'LIF_SEQ_NEXT'
};

/** Rode 1x: instala gatilho onEdit instalável */
function Setup_InstalarAcionador_LIF() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'almox_lif_handleEdit_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('almox_lif_handleEdit_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getActive().toast('Acionador LIF instalado. Edite a COLUNA J dentro de DADOSFRACIONAMENTOS.');
}

/** Handler do sequencial */
function almox_lif_handleEdit_(e) {
  const ss = (e && e.source) ? e.source : SpreadsheetApp.getActive();
  const rngNomeado = ss.getRangeByName(CFG_LIF.NOME_INTERVALO);
  if (!rngNomeado || !e || !e.range) return;

  const shAlvo = rngNomeado.getSheet();
  const shEdit = e.range.getSheet();
  if (shEdit.getSheetId() !== shAlvo.getSheetId()) return;

  const editFirstCol = e.range.getColumn();
  const editLastCol  = editFirstCol + e.range.getNumColumns() - 1;
  if (CFG_LIF.COL_EDIT < editFirstCol || CFG_LIF.COL_EDIT > editLastCol) return;

  const topo = rngNomeado.getRow();
  const base = rngNomeado.getLastRow();
  const startRow = Math.max(e.range.getRow(), topo);
  const endRow   = Math.min(e.range.getLastRow(), base);
  if (endRow < startRow) return;

  const numRows = endRow - startRow + 1;
  const valsJ = shAlvo.getRange(startRow, CFG_LIF.COL_EDIT,  numRows, 1).getValues();
  const valsP = shAlvo.getRange(startRow, CFG_LIF.COL_SAIDA, numRows, 1).getValues();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    almox_lif_ensureCounterInitialized_(ss, rngNomeado);
    let next = almox_lif_getCounter_();

    const out = valsP.map(r => [r[0]]);
    let mudou = false;

    for (let i = 0; i < numRows; i++) {
      const vJ = almox_lif_toText_(valsJ[i][0]).trim();
      const vP = almox_lif_toText_(valsP[i][0]).trim();

      if (vJ === '') {
        if (vP !== '') { out[i][0] = ''; mudou = true; }
      } else {
        if (vP === '') {
          out[i][0] = CFG_LIF.PREFIXO + String(next).padStart(CFG_LIF.DIGITOS, '0');
          next++;
          mudou = true;
        }
      }
    }

    if (!mudou) return;

    const rangeOut = shAlvo.getRange(startRow, CFG_LIF.COL_SAIDA, numRows, 1);

    // ✅ Escrita via script SEMPRE permitida + snapshot atualizado só desse range
    almox_anti_runUnlocked_(ss, () => {
      rangeOut.setValues(out);
    }, [rangeOut]);

    almox_lif_setCounter_(next);
  } finally {
    lock.releaseLock();
  }
}

function almox_lif_ensureCounterInitialized_(ss, rngNomeado) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(CFG_LIF.PROP_KEY);
  if (raw && Number(raw) > 0) return;

  const sh = rngNomeado.getSheet();
  const startRow = rngNomeado.getRow();
  const numRows  = rngNomeado.getNumRows();

  const valsP = sh.getRange(startRow, CFG_LIF.COL_SAIDA, numRows, 1).getValues();
  const re = new RegExp('^' + almox_lif_escapeRegex_(CFG_LIF.PREFIXO) + '(\\d{' + CFG_LIF.DIGITOS + '})$');

  let maxNum = 0;
  for (let i = 0; i < valsP.length; i++) {
    const t = almox_lif_toText_(valsP[i][0]).trim();
    const m = t.match(re);
    if (m) maxNum = Math.max(maxNum, Number(m[1]));
  }
  almox_lif_setCounter_(maxNum + 1);
}
function almox_lif_getCounter_() {
  const raw = PropertiesService.getDocumentProperties().getProperty(CFG_LIF.PROP_KEY);
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? n : 1;
}
function almox_lif_setCounter_(n) {
  PropertiesService.getDocumentProperties().setProperty(CFG_LIF.PROP_KEY, String(Math.max(1, Math.floor(n))));
}
function almox_lif_toText_(v) { return v == null ? '' : String(v); }
function almox_lif_escapeRegex_(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/* =====================================================================
 * 2) ETIQUETAS (Slides/PDF)
 * ===================================================================== */
const CFG_ETQ = {
  template: { slidesId: '1xfAGxKz816Z7JPICPVC3JPZVGvM2UV5HphbpbsYV1QY' }, // fallback
  text: { fontFamily: 'Arial', fontSizePt: 5.4, lineSpacing: 1.1, avgCharEm: 0.55, color: '#000000' },
  barcode: { dpi: 96, codeType: 'Code128' },
  data: { namedRange: 'ETIQUETAS' },
  output: { tituloBase: 'Etiquetas – 3 por página (absoluto)', salvarNoMesmoPastaDoSheet: true, abrirDialogComLink: true },
  positions: [
    { textX: -2.5, textY: 0, textW: 20, textH: 12, bcX: 4.2, bcY: 21.5, bcW: 20, bcH: 8 },
    { textX: 31,   textY: 0, textW: 20, textH: 12, bcX: 36,  bcY: 21.5, bcW: 20, bcH: 8 },
    { textX: 67,   textY: 0, textW: 20, textH: 12, bcX: 72,  bcY: 21.5, bcW: 20, bcH: 8 }
  ]
};

var __BC_CACHE_ETQ = {};

function Setup_InstalarMenuEtiquetas() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ETQ_addMenu_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('ETQ_addMenu_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getActive().toast('Menu instalado. Reabra a planilha para ver "ACIONADORES".');
}

function ETQ_addMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('ACIONADORES')
    .addItem('▶️ Autorizar / Preparar', 'authorizeEverything_SeqEtq')
    .addSeparator()
    .addItem('🖨️ Impressão (Fracionamentos → Etiquetas → PDF)', 'runFracionamentosToEtiquetasAndPrint_')
    .addItem('🖨️ Impressão (ETIQUETAS direto → PDF)', 'almox_impressaoEtiquetasMulti_')
    .addSeparator()
    .addItem('🔒 Anti Mão-Boba (SETUP)', 'almox_anti_setup')
    .addItem('🔒 Anti Mão-Boba (Aplicar bloqueio)', 'almox_anti_aplicarBloqueio')
    .addItem('🔓 Anti Mão-Boba (Liberar 10 min)', 'almox_anti_liberar10min')
    .addItem('🔒 Anti Mão-Boba (Travar agora)', 'almox_anti_travarAgora')
    .addItem('🧹 Anti Mão-Boba (Remover bloqueio)', 'almox_anti_removerBloqueio')
    .addToUi();
}

function authorizeEverything_SeqEtq() {
  try {
    const id = ETQ__getTemplateSlidesId_();
    try { DriveApp.getFileById(id).getName(); } catch (_) {}
    try { SlidesApp.openById(id); } catch (_) {}
    try { UrlFetchApp.fetch('https://www.gstatic.com/generate_204'); } catch (_) {}
    SpreadsheetApp.getActive().toast('Permissões verificadas/ativadas.', 'ACESSOS', 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast('Autorizar: ' + (e && e.message ? e.message : e), 'ACESSOS', 10);
  }
}

function ETQ__extractSlidesIdFromUrl_(url) {
  if (!url) return '';
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  return m ? m[1] : '';
}
function ETQ__readSlidesUrlFromLiberacao_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LIBERAÇÃO IMPORTRANGE') ||
           ss.getSheets().find(s => s.getName().toLowerCase() === 'liberação importrange' ||
                                    s.getName().toLowerCase() === 'liberacao importrange');
  if (!sh) return '';

  const r = sh.getRange(1, 1);
  let url = '';
  try {
    const rtv = r.getRichTextValue();
    if (rtv) {
      url = rtv.getLinkUrl && rtv.getLinkUrl();
      if (!url && rtv.getRuns) {
        for (const run of rtv.getRuns()) {
          const lk = run.getLinkUrl && run.getLinkUrl();
          if (lk) { url = lk; break; }
        }
      }
    }
  } catch (_) {}
  if (!url) {
    const disp = String(r.getDisplayValue() || '');
    const fm = String(r.getFormula() || '');
    const m1 = disp.match(/https?:\/\/[^\s)"]+/i);
    if (m1) url = m1[0];
    if (!url) {
      const m2 = fm.match(/HYPERLINK\("([^"]+)"/i);
      if (m2) url = m2[1];
    }
  }
  return url || '';
}
function ETQ__getTemplateSlidesId_() {
  const url = ETQ__readSlidesUrlFromLiberacao_();
  const idFromSheet = ETQ__extractSlidesIdFromUrl_(url);
  const fallback = CFG_ETQ.template && CFG_ETQ.template.slidesId ? String(CFG_ETQ.template.slidesId) : '';
  const id = idFromSheet || fallback;
  if (!id) throw new Error('Defina o modelo via hyperlink em "LIBERAÇÃO IMPORTRANGE!A1" ou em CFG_ETQ.template.slidesId.');
  return id;
}

function ETQ__qtyToInt_(raw) {
  if (raw == null || raw === '') return 0;
  let val;
  if (typeof raw === 'number') val = isNaN(raw) ? 0 : raw;
  else {
    let s = String(raw).trim().replace(/\s+/g, '');
    if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.indexOf(',') >= 0) s = s.replace(',', '.');
    val = parseFloat(s); if (isNaN(val)) val = 0;
  }
  let r = Math.round(val);
  if (val > 0 && r === 0) r = 1;
  if (r < 0) r = 0;
  return r;
}

function ETQ_lerListaEtiquetas_() {
  const rng = SpreadsheetApp.getActive().getRangeByName(CFG_ETQ.data.namedRange);
  if (!rng) throw new Error(`Intervalo nomeado não encontrado: ${CFG_ETQ.data.namedRange}`);

  const sh   = rng.getSheet();
  const top  = rng.getRow();
  const left = rng.getColumn();
  const rows = Math.min(200, sh.getMaxRows() - top + 1);
  const cols = Math.max(9, rng.getNumColumns());

  const vals = sh.getRange(top, left, rows, cols).getValues();
  const items = [];
  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];
    const [cod, nome, un, fornecedor, oc, lote, validade, barcodeCell, qtd] = row;
    const q = ETQ__qtyToInt_(qtd);
    if (!String(lote || '').trim() || q <= 0) continue;
    const data = { cod, nome, un, fornecedor, oc, lote, validade, barcodeCell, barcodeFormula: '' };
    for (let k = 0; k < q; k++) items.push(data);
  }
  return items;
}

function almox_impressaoEtiquetasMulti_() {
  ETQ_assertTemplate_();
  ETQ_cleanupOldOutputs_();

  const list = ETQ_lerListaEtiquetas_();
  if (!list || !list.length) {
    SpreadsheetApp.getActive().toast('Nenhuma etiqueta com quantidade > 0 em ETIQUETAS.', 'Impressão', 6);
    return null;
  }

  const pres = ETQ_makeCopyFromTemplate_(`${CFG_ETQ.output.tituloBase} – ${SpreadsheetApp.getActive().getName()} – ${ETQ_agoraFmt_()}`);
  pres.getSlides().forEach(s => s.remove());

  for (let i = 0; i < list.length; i += CFG_ETQ.positions.length) {
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    for (let j = 0; j < CFG_ETQ.positions.length; j++) {
      const item = list[i + j];
      if (!item) break;
      const p = CFG_ETQ.positions[j];

      const textRightMm = p.bcX + p.bcW + 4;
      const textWidthMm = Math.max(0.1, textRightMm - p.textX);

      const bcUrl  = ETQ_barcodeUrlPreferred_(item);
      const bcBlob = ETQ_getBarcodeBlobOnceFromUrl_(bcUrl);

      ETQ_drawText_(slide, p.textX, p.textY, textWidthMm, p.textH, item);
      ETQ_drawBarcode_(slide, p.bcX, p.bcY, p.bcW, p.bcH, bcUrl, bcBlob);
    }
  }

  try { pres.saveAndClose(); } catch(e){}
  try { Utilities.sleep(800); } catch(e){}

  const pdfBlob = DriveApp.getFileById(pres.getId()).getAs('application/pdf');
  const pdfName = `${CFG_ETQ.output.tituloBase} - ${ETQ_agoraCompact_()}.pdf`;
  const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);

  if (CFG_ETQ.output.salvarNoMesmoPastaDoSheet) {
    ETQ_moverArquivoParaMesmaPasta_(pdfFile.getId(), SpreadsheetApp.getActive().getId());
  }

  if (CFG_ETQ.output.abrirDialogComLink) ETQ_mostrarLink_(pdfFile.getUrl());
  else SpreadsheetApp.getActive().toast('PDF gerado: ' + pdfFile.getUrl(), 'Impressão', 8);

  try { DriveApp.getFileById(pres.getId()).setTrashed(true); } catch(e){}
  return pdfFile;
}

/** Fluxo: DADOSFRACIONAMENTOS -> ETIQUETAS -> PDF -> limpa -> nota em H */
function runFracionamentosToEtiquetasAndPrint_() {
  const ss = SpreadsheetApp.getActive();
  const nrFrac = ss.getRangeByName('DADOSFRACIONAMENTOS');
  const nrEtq  = ss.getRangeByName(CFG_ETQ.data.namedRange);
  if (!nrFrac) throw new Error('Named range "DADOSFRACIONAMENTOS" não encontrado.');
  if (!nrEtq)  throw new Error(`Named range "${CFG_ETQ.data.namedRange}" não encontrado.`);

  const shF = nrFrac.getSheet();
  const shE = nrEtq.getSheet();

  const colH = 8;   // quantidade + nota
  const colP = 16;  // lote/código -> ETIQUETAS!F

  const etqTop  = nrEtq.getRow();
  const etqHgt  = nrEtq.getNumRows();
  const etqColF = nrEtq.getColumn() + 5;  // F
  const etqColI = nrEtq.getColumn() + 8;  // I

  const frTop = nrFrac.getRow();
  const frHgt = nrFrac.getNumRows();

  const valsH  = shF.getRange(frTop, colH, frHgt, 1).getValues().map(r => r[0]);
  const notesH = shF.getRange(frTop, colH, frHgt, 1).getNotes().map(r => r[0]);
  const valsP  = shF.getRange(frTop, colP, frHgt, 1).getValues().map(r => r[0]);

  const toF = [];
  const toI = [];
  const rowsFracAbs = [];

  for (let i = 0; i < frHgt; i++) {
    const noteEmpty  = String(notesH[i] || '').trim() === '';
    const loteFilled = String(valsP[i]   || '').trim() !== '';
    if (!noteEmpty || !loteFilled) continue;

    const q = ETQ__qtyToInt_(valsH[i]);
    if (q > 0) {
      toF.push([ valsP[i] ]);
      toI.push([ q ]);
      rowsFracAbs.push(frTop + i);
    }
  }

  if (!rowsFracAbs.length) {
    SpreadsheetApp.getActive().toast('Nada para imprimir: sem linhas elegíveis.', 'Impressão', 6);
    return;
  }

  // primeira linha vazia em F dentro do named range
  const colF_disp = shE.getRange(etqTop, etqColF, etqHgt, 1).getDisplayValues().map(r => String(r[0] || '').trim());
  const firstEmptyIdx = colF_disp.findIndex(v => v === '');
  let writeR = (firstEmptyIdx >= 0) ? (etqTop + firstEmptyIdx) : (etqTop + etqHgt);

  const n = toF.length;
  const needLast = writeR + n - 1;
  if (needLast > shE.getMaxRows()) shE.insertRowsAfter(shE.getMaxRows(), needLast - shE.getMaxRows());

  const rangeF = shE.getRange(writeR, etqColF, n, 1);
  const rangeI = shE.getRange(writeR, etqColI, n, 1);

  // 1) Escreve em ETIQUETAS com permissão garantida + snapshot pontual
  almox_anti_runUnlocked_(ss, () => {
    rangeF.setValues(toF);
    rangeI.setValues(toI);
  }, [rangeF, rangeI]);

  // 2) Gera PDF (não precisa ficar em modo edit liberado)
  try {
    almox_impressaoEtiquetasMulti_();
  } finally {
    // 3) Limpa ETIQUETAS com permissão garantida + snapshot pontual
    almox_anti_runUnlocked_(ss, () => {
      rangeF.clearContent();
      rangeI.clearContent();
    }, [rangeF, rangeI]);
  }

  // Carimba NOTA em H nas linhas corretas (contíguas em blocos)
  const stamp = (function(){
    const d=new Date(),p=x=>x<10?'0'+x:x;
    return `Impresso, ${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  })();

  const groups = _groupContiguousRows_(rowsFracAbs);
  groups.forEach(g => {
    const notesMatrix = Array(g.len).fill([stamp]);
    // nota não afeta DV, mas ainda assim é uma escrita — e pode estar em range travado
    const rgNotes = shF.getRange(g.start, colH, g.len, 1);

    almox_anti_runUnlocked_(ss, () => {
      rgNotes.setNotes(notesMatrix);
    }, []); // sem necessidade de snapshot (notes não entram na validação)
  });

  SpreadsheetApp.getActive().toast('Impressão concluída ✅', 'Impressão', 6);
}

function _groupContiguousRows_(rowsAbs) {
  const rows = Array.from(new Set(rowsAbs)).sort((a,b)=>a-b);
  if (!rows.length) return [];
  const out = [];
  let start = rows[0], prev = rows[0], len = 1;
  for (let i=1;i<rows.length;i++){
    const r = rows[i];
    if (r === prev + 1) { len++; prev = r; continue; }
    out.push({ start, len });
    start = r; prev = r; len = 1;
  }
  out.push({ start, len });
  return out;
}

/* -------- desenho / barcode / util ETQ -------- */
function ETQ_drawText_(slide, xMm, yMm, wMm, hMm, data) {
  const box = slide.insertTextBox('', ETQ_mm2pt_(xMm), ETQ_mm2pt_(yMm), ETQ_mm2pt_(wMm), ETQ_mm2pt_(hMm));
  try { box.getAutofit().setAutofitType(SlidesApp.AutofitType.NONE); } catch (e) {}

  const l1 = ETQ_lineClampNoWrap_('NOME:',  ETQ_n_(data.nome),       wMm);
  const l2 = ETQ_lineClampNoWrap_('UN.:',   ETQ_n_(data.un),         wMm);
  const l3 = ETQ_lineClampNoWrap_('END:',   ETQ_n_(data.fornecedor), wMm);
  const l4 = ETQ_lineClampNoWrap_('Nº OC:', ETQ_n_(data.oc),         wMm);
  const l5 = ETQ_lineClampNoWrap_('LOTE:',  ETQ_n_(data.lote),       wMm);

  const NBSP = '\u00A0';
  const l6Raw = `VALIDADE:${NBSP}${ETQ_fmtValidade_(data.validade)}${NBSP}${NBSP}COD:${NBSP}${ETQ_n_(data.cod)}`;
  const l6 = ETQ_lineClampRawNoWrap_(l6Raw, wMm);

  const text = box.getText();
  text.setText([l1, l2, l3, l4, l5, l6].join('\n'));

  const style = text.getTextStyle();
  style.setFontFamily(CFG_ETQ.text.fontFamily)
       .setFontSize(CFG_ETQ.text.fontSizePt)
       .setBold(true)
       .setForegroundColor(CFG_ETQ.text.color);

  try { text.getParagraphStyle().setLineSpacing(CFG_ETQ.text.lineSpacing); } catch (e) {}
  box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
}
function ETQ_drawBarcode_(slide, xMm, yMm, wMm, hMm, url, blob) {
  try {
    let img;
    try { img = slide.insertImage(url); } catch (e) {}
    if (!img && blob) { try { img = slide.insertImage(blob); } catch (e) {} }
    if (!img) throw new Error('Sem imagem de barcode');

    img.setLeft(ETQ_mm2pt_(xMm))
       .setTop(ETQ_mm2pt_(yMm))
       .setWidth(ETQ_mm2pt_(wMm))
       .setHeight(ETQ_mm2pt_(hMm));
  } catch (e) {
    const warn = slide.insertTextBox('[BARCODE INDISPONÍVEL]', ETQ_mm2pt_(xMm), ETQ_mm2pt_(yMm), ETQ_mm2pt_(wMm), ETQ_mm2pt_(hMm));
    warn.getText().getTextStyle().setFontFamily('Arial').setFontSize(8).setBold(true).setForegroundColor('#AA0000');
    warn.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  }
}
function ETQ_barcodeUrlFromCod_(cod) {
  const data = encodeURIComponent(String(cod || ''));
  return `https://barcode.tec-it.com/barcode.ashx?data=${data}&code=${CFG_ETQ.barcode.codeType}&dpi=${CFG_ETQ.barcode.dpi}`;
}
function ETQ_barcodeUrlPreferred_(d) {
  if (d && d.cod != null && String(d.cod).trim() !== '') return ETQ_barcodeUrlFromCod_(d.cod);
  const val = String(d.barcodeCell || '').trim();
  if (/^https?:\/\//i.test(val)) return val;
  return ETQ_barcodeUrlFromCod_(d.cod);
}
function ETQ_getBarcodeBlobOnceFromUrl_(url) {
  if (!url) return null;
  try {
    let blob = __BC_CACHE_ETQ[url];
    if (!blob) {
      for (let i = 0; i < 3; i++) {
        try { blob = UrlFetchApp.fetch(url).getBlob().setName('barcode.png'); break; }
        catch (e) { Utilities.sleep(400 * (i + 1)); }
      }
      if (blob) __BC_CACHE_ETQ[url] = blob;
    }
    return blob || null;
  } catch (e) { return null; }
}
function ETQ_mm2pt_(mm) { return mm * 2.834645669291339; }
function ETQ_agoraFmt_(){ const d=new Date(),p=n=>n<10?'0'+n:n; return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; }
function ETQ_agoraCompact_(){ const d=new Date(),p=n=>n<10?'0'+n:n; return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`; }
function ETQ_n_(x){ return (x==null||x==='') ? '' : String(x); }
function ETQ_formatDateBR_(val){
  if (val == null || val === '') return '';
  const s = String(val).trim();
  if (val instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s)){
    try{ const d = (val instanceof Date) ? val : new Date(val); if (isNaN(d)) return s;
      const p=n=>n<10?'0'+n:n; return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()}`; }catch(e){ return s; }
  }
  if (/^\d+(?:[\\.,]\d+)?$/.test(s)) return s;
  return s;
}
function ETQ_fmtValidade_(v){ return ETQ_formatDateBR_(v); }
function ETQ_makeCopyFromTemplate_(title){
  const slidesId = ETQ__getTemplateSlidesId_();
  const src  = DriveApp.getFileById(slidesId);
  const file = src.makeCopy(title);
  return SlidesApp.openById(file.getId());
}
function ETQ_mostrarLink_(url){
  const html = HtmlService.createHtmlOutput(
    `<div style="font:14px Arial;padding:16px;">
      <div><b>PDF gerado (3 etiquetas por página, coordenadas absolutas).</b></div>
      <div style="margin-top:8px"><a target="_blank" href="${url}">Abrir PDF</a></div>
      <div style="margin-top:12px;color:#666">Imprima em 100% (tamanho real). Papel 102×25 mm. Sensor Gap.</div>
    </div>`).setWidth(380).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Impressão');
}
function ETQ_assertTemplate_(){ ETQ__getTemplateSlidesId_(); }
function ETQ_moverArquivoParaMesmaPasta_(fileId, sheetId){
  try{
    const file   = DriveApp.getFileById(fileId);
    const ssFile = DriveApp.getFileById(sheetId);
    const parents = ssFile.getParents();
    if (parents.hasNext()){
      const folder = parents.next();
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
  } catch (e) {}
}
function ETQ_cleanupOldOutputs_(){
  try{
    const prefix = CFG_ETQ.output.tituloBase;
    const ssId = SpreadsheetApp.getActive().getId();
    const ssFile = DriveApp.getFileById(ssId);
    const parents = ssFile.getParents();
    const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const folderId = folder.getId();

    const esc = s => String(s).replace(/'/g, "\\'");
    const q =
      "trashed = false and '" + folderId + "' in parents and " +
      "(mimeType = 'application/pdf' or mimeType = 'application/vnd.google-apps.presentation') and " +
      "title contains '" + esc(prefix) + "'";

    let it;
    try { it = DriveApp.searchFiles(q); }
    catch (e) { it = folder.getFiles(); }

    const keep = new Set([String(CFG_ETQ.template.slidesId || ''), ssId]);

    while (it.hasNext()){
      const f = it.next();
      try {
        const name = f.getName();
        const mime = f.getMimeType();
        const isTarget = String(name).indexOf(prefix) !== -1 &&
                         (mime === 'application/pdf' || mime === 'application/vnd.google-apps.presentation');
        if (isTarget && !keep.has(f.getId())) f.setTrashed(true);
      } catch (e2) {}
    }
  } catch(e){}
}
function ETQ_charsCapacity_(wMm) {
  const mmPerPt = 0.352777778;
  const fontPt = (CFG_ETQ.text && CFG_ETQ.text.fontSizePt) || 5.4;
  const avgEm  = (CFG_ETQ.text && CFG_ETQ.text.avgCharEm)  || 0.55;
  const charMm = fontPt * avgEm * mmPerPt;
  return Math.max(1, Math.floor(wMm / Math.max(0.1, charMm)));
}
function ETQ_lineClampNoWrap_(label, value, wMm) {
  const NBSP = '\u00A0';
  const lab = String(label || '').replace(/ /g, NBSP);
  const val = ETQ_n_(value).replace(/ /g, NBSP);
  const full = lab + NBSP + val;
  const cap = ETQ_charsCapacity_(wMm);
  return (full.length <= cap) ? full : full.slice(0, cap);
}
function ETQ_lineClampRawNoWrap_(raw, wMm) {
  const NBSP = '\u00A0';
  const s = String(raw || '').replace(/ /g, NBSP);
  const cap = ETQ_charsCapacity_(wMm);
  return (s.length <= cap) ? s : s.slice(0, cap);
}

/* =====================================================================
 * 3) ANTI MÃO-BOBA — SETUP / APLICAR / LIBERAR / TRAVAR / REMOVER
 * ===================================================================== */

function almox_anti_setup() {
  const ss = SpreadsheetApp.getActive();

  // Config sheet
  let sh = ss.getSheetByName(ALMOX_ANTI.CONFIG_SHEET);
  if (!sh) sh = ss.insertSheet(ALMOX_ANTI.CONFIG_SHEET);

  sh.getRange(ALMOX_ANTI.EDIT_MODE_A1).setValue(false);
  try { sh.hideSheet(); } catch (_) {}

  const cell = sh.getRange(ALMOX_ANTI.EDIT_MODE_A1);
  const existing = ss.getNamedRanges().find(nr => nr.getName() === ALMOX_ANTI.EDIT_MODE_NAMED);
  if (existing) existing.setRange(cell);
  else ss.setNamedRange(ALMOX_ANTI.EDIT_MODE_NAMED, cell);

  // Trigger diário 03:00
  almox_anti_installDailyTrigger_();

  ss.toast('Anti Mão-Boba: setup concluído ✅', 'Anti Mão-Boba', 4);
}

/**
 * Aplica bloqueio:
 * - Copia snapshot dos blocos protegidos (sparse)
 * - Aplica DataValidation (bloqueio) somente nessas áreas
 * - Respeita "exceto algumas células"
 */
function almox_anti_aplicarBloqueio(silent) {
  const SILENT = !!silent;
  const ss = SpreadsheetApp.getActive();

  // garante trigger diário
  almox_anti_installDailyTrigger_();

  const modeCell = ss.getRangeByName(ALMOX_ANTI.EDIT_MODE_NAMED);
  if (!modeCell) {
    if (!SILENT) SpreadsheetApp.getUi().alert('Anti Mão-Boba', 'Rode primeiro: "almox_anti_setup()".', SpreadsheetApp.getUi().ButtonSet.OK);
    console.error('Anti Mão-Boba: faltou rodar almox_anti_setup().');
    return;
  }

  const sheetProtections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const rangeProtections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  if (!sheetProtections.length && !rangeProtections.length) {
    if (!SILENT) SpreadsheetApp.getUi().alert('Anti Mão-Boba', 'Nenhuma proteção (ABA ou RANGE) foi encontrada.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const snapCache = {};
  function getSnapFor_(sh) {
    const id = sh.getSheetId();
    if (!snapCache[id]) snapCache[id] = almox_anti_getSnapSheet_(ss, sh, { reset: true, createIfMissing: true });
    return snapCache[id];
  }

  let blocksApplied = 0;

  // Proteção de ABA
  sheetProtections.forEach(p => {
    const sh = p.getRange().getSheet();
    const snap = getSnapFor_(sh);

    const base = almox_anti_getEffectiveBaseRange_(sh, p);
    const protectedRects = almox_anti_computeProtectedRectsFromSheetProtection_(p, base);

    // snapshot sparse
    almox_anti_writeSnapshotSparse_(snap, sh, protectedRects);

    // aplica DV por retângulo
    const rule = almox_anti_buildRuleForSnap_(snap.getName());
    protectedRects.forEach(r => {
      const rg = sh.getRange(r.sr, r.sc, r.er - r.sr + 1, r.ec - r.sc + 1);
      almox_anti_retry_(() => rg.setDataValidation(rule), 'setDataValidation.rect');
      blocksApplied++;
    });

    // remove DV do Anti nas exceções
    const unprot = (p.getUnprotectedRanges && p.getUnprotectedRanges()) || [];
    almox_anti_removeOnlyOurValidation_(unprot);
  });

  // Proteção de RANGE
  rangeProtections.forEach(p => {
    const rg = p.getRange();
    const sh = rg.getSheet();
    const snap = getSnapFor_(sh);

    const rect = { sr: rg.getRow(), sc: rg.getColumn(), er: rg.getLastRow(), ec: rg.getLastColumn() };
    almox_anti_writeSnapshotSparse_(snap, sh, [rect]);

    const rule = almox_anti_buildRuleForSnap_(snap.getName());
    almox_anti_retry_(() => rg.setDataValidation(rule), 'setDataValidation.range');
    blocksApplied++;
  });

  if (!SILENT) ss.toast(`Anti Mão-Boba aplicado ✅ blocos/ranges: ${blocksApplied}.`, 'Anti Mão-Boba', 8);
}

function almox_anti_liberar10min() {
  const ss = SpreadsheetApp.getActive();
  const cell = ss.getRangeByName(ALMOX_ANTI.EDIT_MODE_NAMED);
  if (!cell) throw new Error('ANTI_EDIT_MODE não encontrado. Rode almox_anti_setup().');

  cell.setValue(true);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'almox_anti_travarAgora')
    .forEach(t => { try { ScriptApp.deleteTrigger(t); } catch (_) {} });

  ScriptApp.newTrigger('almox_anti_travarAgora')
    .timeBased()
    .after(ALMOX_ANTI.UNLOCK_MINUTES * 60 * 1000)
    .create();

  ss.toast(`Edição liberada por ${ALMOX_ANTI.UNLOCK_MINUTES} min.`, 'Anti Mão-Boba', 5);
}

function almox_anti_travarAgora() {
  const ss = SpreadsheetApp.getActive();
  const cell = ss.getRangeByName(ALMOX_ANTI.EDIT_MODE_NAMED);
  if (cell) cell.setValue(false);

  // Atualiza snapshots ao travar (reduz "inválido")
  almox_anti_refreshAllSnapshots_();

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'almox_anti_travarAgora')
    .forEach(t => { try { ScriptApp.deleteTrigger(t); } catch (_) {} });

  ss.toast('Edição travada ✅', 'Anti Mão-Boba', 3);
}

function almox_anti_removerBloqueio() {
  const ss = SpreadsheetApp.getActive();
  let removed = 0;

  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
    const sh = p.getRange().getSheet();
    const base = almox_anti_getEffectiveBaseRange_(sh, p);
    const protectedRects = almox_anti_computeProtectedRectsFromSheetProtection_(p, base);

    protectedRects.forEach(r => {
      const rg = sh.getRange(r.sr, r.sc, r.er - r.sr + 1, r.ec - r.sc + 1);
      const dv = rg.getCell(1, 1).getDataValidation();
      if (dv && almox_anti_isOurDv_(dv)) {
        rg.clearDataValidations();
        removed++;
      }
    });

    const unprot = (p.getUnprotectedRanges && p.getUnprotectedRanges()) || [];
    almox_anti_removeOnlyOurValidation_(unprot);
  });

  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    const rg = p.getRange();
    const dv = rg.getCell(1, 1).getDataValidation();
    if (dv && almox_anti_isOurDv_(dv)) {
      rg.clearDataValidations();
      removed++;
    }
  });

  ss.toast(`Anti Mão-Boba removido em ${removed} bloco(s)/range(s).`, 'Anti Mão-Boba', 6);
}

/* -------- helpers ANTI -------- */

function almox_anti_getSnapSheet_(ss, sheet, opts) {
  opts = opts || {};
  const name = ALMOX_ANTI.SNAP_PREFIX + sheet.getSheetId();
  let snap = ss.getSheetByName(name);

  if (opts.reset) {
    try {
      const active = ss.getActiveSheet();
      if (snap) {
        if (active && active.getSheetId && active.getSheetId() === snap.getSheetId()) {
          ss.setActiveSheet(sheet);
        }
        ss.deleteSheet(snap);
      }
    } catch (_) {}
    snap = null;
  }

  if (!snap && opts.createIfMissing) snap = ss.insertSheet(name);
  if (snap) { try { snap.hideSheet(); } catch (_) {} }
  return snap;
}

function almox_anti_buildRuleForSnap_(snapSheetName) {
  // Regra: válido se
  //  - ANTI_EDIT_MODE ligado (scripts / liberação temporária)
  //  - OU valor atual == snapshot
  const f =
    `=OR(${ALMOX_ANTI.EDIT_MODE_NAMED};` +
    `INDIRECT(ADDRESS(ROW();COLUMN();4))=` +
    `INDIRECT("'${snapSheetName}'!"&ADDRESS(ROW();COLUMN();4)))`;

  return SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(f)
    .setAllowInvalid(false)
    .setHelpText(ALMOX_ANTI.HELP_TEXT)
    .build();
}

function almox_anti_isOurDv_(dv) {
  try {
    const help = dv.getHelpText && dv.getHelpText();
    const allow = dv.getAllowInvalid && dv.getAllowInvalid();
    return help === ALMOX_ANTI.HELP_TEXT && allow === false;
  } catch (_) { return false; }
}

function almox_anti_removeOnlyOurValidation_(ranges) {
  for (const rg of ranges || []) {
    if (!rg) continue;
    const dvs = rg.getDataValidations();
    let changed = false;
    for (let r = 0; r < dvs.length; r++) {
      for (let c = 0; c < dvs[0].length; c++) {
        const dv = dvs[r][c];
        if (dv && almox_anti_isOurDv_(dv)) { dvs[r][c] = null; changed = true; }
      }
    }
    if (changed) rg.setDataValidations(dvs);
  }
}

function almox_anti_getEffectiveBaseRange_(sheet, sheetProtection) {
  let lastRow = Math.max(1, sheet.getLastRow());
  let lastCol = Math.max(1, sheet.getLastColumn());

  const unprot = (sheetProtection.getUnprotectedRanges && sheetProtection.getUnprotectedRanges()) || [];
  unprot.forEach(rg => {
    if (!rg) return;
    if (rg.getSheet().getSheetId() !== sheet.getSheetId()) return;
    lastRow = Math.max(lastRow, rg.getLastRow());
    lastCol = Math.max(lastCol, rg.getLastColumn());
  });

  lastRow = Math.min(lastRow, ALMOX_ANTI.SAFE_MAX_ROW);
  lastCol = Math.min(lastCol, ALMOX_ANTI.SAFE_MAX_COL);

  return sheet.getRange(1, 1, lastRow, lastCol);
}

function almox_anti_computeProtectedRectsFromSheetProtection_(sheetProtection, baseRange) {
  const sh = baseRange.getSheet();

  const bSr = baseRange.getRow();
  const bSc = baseRange.getColumn();
  const bEr = bSr + baseRange.getNumRows() - 1;
  const bEc = bSc + baseRange.getNumColumns() - 1;

  const unprot = (sheetProtection.getUnprotectedRanges && sheetProtection.getUnprotectedRanges()) || [];
  const rects = [];

  unprot.forEach(rg => {
    if (!rg) return;
    if (rg.getSheet().getSheetId() !== sh.getSheetId()) return;

    const sr = Math.max(bSr, rg.getRow());
    const sc = Math.max(bSc, rg.getColumn());
    const er = Math.min(bEr, rg.getLastRow());
    const ec = Math.min(bEc, rg.getLastColumn());
    if (sr > er || sc > ec) return;

    rects.push({ sr, sc, er, ec });
  });

  if (!rects.length) return [{ sr: bSr, sc: bSc, er: bEr, ec: bEc }];

  const rowCuts = new Set([bSr, bEr + 1]);
  const colCuts = new Set([bSc, bEc + 1]);
  rects.forEach(r => { rowCuts.add(r.sr); rowCuts.add(r.er + 1); colCuts.add(r.sc); colCuts.add(r.ec + 1); });

  const rows = Array.from(rowCuts).sort((a, b) => a - b);
  const cols = Array.from(colCuts).sort((a, b) => a - b);

  const protectedRects = [];

  for (let i = 0; i < rows.length - 1; i++) {
    const sr = rows[i], er = rows[i + 1] - 1;
    if (sr > er) continue;

    for (let j = 0; j < cols.length - 1; j++) {
      const sc = cols[j], ec = cols[j + 1] - 1;
      if (sc > ec) continue;

      let isUnprotected = false;
      for (const u of rects) {
        if (sr >= u.sr && er <= u.er && sc >= u.sc && ec <= u.ec) { isUnprotected = true; break; }
      }
      if (!isUnprotected) protectedRects.push({ sr, sc, er, ec });
    }
  }

  return protectedRects;
}

function almox_anti_writeSnapshotSparse_(snapSheet, srcSheet, rects) {
  if (!rects || !rects.length) return;

  let maxR = 1, maxC = 1;
  rects.forEach(r => { maxR = Math.max(maxR, r.er); maxC = Math.max(maxC, r.ec); });
  almox_anti_ensureSize_(snapSheet, maxR, maxC);

  for (const r of rects) {
    const nr = r.er - r.sr + 1;
    const nc = r.ec - r.sc + 1;

    for (let rOff = 0; rOff < nr; rOff += ALMOX_ANTI.ROW_CHUNK) {
      const h = Math.min(ALMOX_ANTI.ROW_CHUNK, nr - rOff);

      for (let cOff = 0; cOff < nc; cOff += ALMOX_ANTI.COL_CHUNK) {
        const w = Math.min(ALMOX_ANTI.COL_CHUNK, nc - cOff);

        const src = srcSheet.getRange(r.sr + rOff, r.sc + cOff, h, w);
        const dst = snapSheet.getRange(r.sr + rOff, r.sc + cOff, h, w);

        almox_anti_retry_(() => {
          src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        }, 'snapshot.sparse.copyTo');
      }
      Utilities.sleep(10);
    }
  }
}

function almox_anti_ensureSize_(sheet, minRows, minCols) {
  const r = sheet.getMaxRows();
  const c = sheet.getMaxColumns();
  const addR = Math.max(0, minRows - r);
  const addC = Math.max(0, minCols - c);

  if (addR > 20000 || addC > 1000) {
    throw new Error(`Anti Mão-Boba: snapshot precisaria crescer demais (rows +${addR}, cols +${addC}).`);
  }

  if (addR > 0) almox_anti_retry_(() => sheet.insertRowsAfter(r, addR), 'insertRowsAfter');
  if (addC > 0) almox_anti_retry_(() => sheet.insertColumnsAfter(c, addC), 'insertColumnsAfter');
}

function almox_anti_retry_(fn, label) {
  let lastErr = null;
  for (let i = 0; i < ALMOX_ANTI.RETRIES; i++) {
    try { return fn(); }
    catch (e) {
      lastErr = e;
      const msg = String(e && e.message ? e.message : e);
      if (/Service error|Spreadsheets|Internal error|Limit exceeded/i.test(msg)) {
        Utilities.sleep(350 * (i + 1));
        continue;
      }
      throw e;
    }
  }
  throw new Error(`${label || 'operação'} falhou após ${ALMOX_ANTI.RETRIES} tentativas: ${
    lastErr && lastErr.message ? lastErr.message : lastErr
  }`);
}

/** Atualiza snapshots de todas as proteções (use com parcimônia; rotina diária usa isso). */
function almox_anti_refreshAllSnapshots_() {
  const ss = SpreadsheetApp.getActive();
  const snapCache = {};

  function getSnapFor_(sh) {
    const id = sh.getSheetId();
    if (!snapCache[id]) snapCache[id] = almox_anti_getSnapSheet_(ss, sh, { reset: true, createIfMissing: true });
    return snapCache[id];
  }

  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
    const sh = p.getRange().getSheet();
    const snap = getSnapFor_(sh);
    const base = almox_anti_getEffectiveBaseRange_(sh, p);
    const protectedRects = almox_anti_computeProtectedRectsFromSheetProtection_(p, base);
    almox_anti_writeSnapshotSparse_(snap, sh, protectedRects);

    const unprot = (p.getUnprotectedRanges && p.getUnprotectedRanges()) || [];
    almox_anti_removeOnlyOurValidation_(unprot);
  });

  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    const rg = p.getRange();
    const sh = rg.getSheet();
    const snap = getSnapFor_(sh);
    const rect = { sr: rg.getRow(), sc: rg.getColumn(), er: rg.getLastRow(), ec: rg.getLastColumn() };
    almox_anti_writeSnapshotSparse_(snap, sh, [rect]);
  });
}
