/*********************************************************
 *  ALMOX — FRACIONAMENTO + ETIQUETAS + ANTI-EDIT
 *
 *  ✅ SEM "CONFIRMAÇÃO DE COMANDOS" / SEM rodar como proprietário
 *  ✅ Sequencial LIF em DADOSFRACIONAMENTOS (col J -> escreve em P)
 *  ✅ Impressão (ETIQUETAS) + PDF + limpeza + nota em H
 *  ✅ Anti-edit: bloqueia o proprietário nas áreas protegidas
 *  ✅ Scripts SEMPRE podem escrever (sem virar "inválido" depois)
 *********************************************************/

/* ===================== [ANTI-EDIT] ====================== */
const ANTI_SHEET_NAME = '_ANTI_INDENTIFICACAO_';
const ANTI_MODE_CELL = 'A1';
const ANTI_MODE_NAMED_RANGE = 'ANTI_EDIT_MODE';
const ANTI_SNAP_PREFIX = '_ANTI_SNAP_';
const ANTI_CHUNK_ROWS_ = 200;

/* ========= Menu ========= */
function onOpen() {
  buildMenu_();
}

function ETQ_addMenu_() {
  buildMenu_();
}

function buildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('ACIONADORES')
    .addItem("🖨️ Impressão (ETIQUETAS direto → PDF)'", 'almox_impressaoEtiquetasMulti_')
    .addItem('Liberar permissões', 'liberarPermissoes')
    .addItem('Anti-edit', 'antiEdit_')
    .addToUi();
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
    withAntiEditTemporaryUnblock_([rangeOut], () => {
      rangeOut.setValues(out);
    });

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

// (menu criado por onOpen/ETQ_addMenu_)

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

function liberarPermissoes() {
  try { SpreadsheetApp.getActive(); } catch (_) {}
  try { SpreadsheetApp.getUi(); } catch (_) {}
  try { SpreadsheetApp.getActive().getId(); } catch (_) {}
  try { Session.getActiveUser().getEmail(); } catch (_) {}
  try { PropertiesService.getDocumentProperties().getKeys(); } catch (_) {}
  try { ScriptApp.getProjectTriggers(); } catch (_) {}
  try { UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); } catch (_) {}
  try { MailApp.getRemainingDailyQuota(); } catch (_) {}
  try { SpreadsheetApp.getActive().toast('Permissões solicitadas com sucesso.', 'Permissões', 6); } catch (_) {}
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
  withAntiEditTemporaryUnblock_([rangeF, rangeI], () => {
    rangeF.setValues(toF);
    rangeI.setValues(toI);
  });

  // 2) Gera PDF (não precisa ficar em modo edit liberado)
  try {
    almox_impressaoEtiquetasMulti_();
  } finally {
    // 3) Limpa ETIQUETAS com permissão garantida + snapshot pontual
    withAntiEditTemporaryUnblock_([rangeF, rangeI], () => {
      rangeF.clearContent();
      rangeI.clearContent();
    });
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

    withAntiEditTemporaryUnblock_([rgNotes], () => {
      rgNotes.setNotes(notesMatrix);
    });
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
 * 3) ANTI-EDIT (snapshot + validação com rejeição)
 * ===================================================================== */

function antiEditIdentify_() {
  ensureAntiEditSetup_();
  SpreadsheetApp.getActive().toast('Anti-edit identificado.', 'Anti-edit', 4);
}

function antiEdit_() {
  const ss = SpreadsheetApp.getActive();
  deleteAntiEditSheets_(ss);
  antiEditRemove_();
  antiEditIdentify_();
  antiEditApply_();
  ss.toast('Anti-edit concluído.', 'Anti-edit', 6);
}

function antiEditApply_() {
  const ss = SpreadsheetApp.getActive();
  ensureAntiEditSetupForSpreadsheet_(ss);

  const ruleBySnap = {};
  const sheetProtections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
  const rangeProtections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];

  sheetProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const sheet = p.getRange().getSheet();
    const unprotected = p.getUnprotectedRanges ? (p.getUnprotectedRanges() || []) : [];
    const baseRange = buildBaseRangeWithExceptions_(sheet, unprotected);
    if (!baseRange) return;

    const snapSheet = ensureAntiSnapSheet_(sheet, baseRange);
    writeSnapshotRange_(snapSheet, baseRange);

    const snapName = snapSheet.getName();
    const rule = ruleBySnap[snapName] || (ruleBySnap[snapName] = buildAntiValidationRule_(snapName));
    applyAntiValidationToRangeFast_(baseRange, rule);
    clearAntiValidationForRanges_(unprotected);
  });

  rangeProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const range = p.getRange();
    if (!range) return;

    const snapSheet = ensureAntiSnapSheet_(range.getSheet(), range);
    writeSnapshotRange_(snapSheet, range);

    const snapName = snapSheet.getName();
    const rule = ruleBySnap[snapName] || (ruleBySnap[snapName] = buildAntiValidationRule_(snapName));
    applyAntiValidationToRangeFast_(range, rule);
  });

  ss.toast('Anti-edit aplicado.', 'Anti-edit', 5);
}

function antiEditRemove_() {
  const ss = SpreadsheetApp.getActive();
  const sheetProtections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
  const rangeProtections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];

  sheetProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const sheet = p.getRange().getSheet();
    const unprotected = p.getUnprotectedRanges ? (p.getUnprotectedRanges() || []) : [];
    const baseRange = buildBaseRangeWithExceptions_(sheet, unprotected);
    if (!baseRange) return;
    clearAntiValidationInRangeFast_(baseRange);
    clearAntiValidationForRanges_(unprotected);
  });

  rangeProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const range = p.getRange();
    if (range) clearAntiValidationInRangeFast_(range);
  });

  ss.toast('Anti-edit removido.', 'Anti-edit', 5);
}

function deleteAntiEditSheets_(ss) {
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf('_ANTI') === 0 || name.indexOf('_Anti') === 0) {
      try { ss.deleteSheet(sheet); } catch (_) {}
    }
  });
}

function ensureAntiEditSetup_() {
  return ensureAntiEditSetupForSpreadsheet_(SpreadsheetApp.getActive());
}

function ensureAntiEditSetupForSpreadsheet_(ss) {
  let sheet = ss.getSheetByName(ANTI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANTI_SHEET_NAME);
    sheet.hideSheet();
  }

  const modeCell = sheet.getRange(ANTI_MODE_CELL);
  modeCell.setValue(false);
  sheet.hideSheet();

  const existing = ss.getNamedRanges().find(nr => nr.getName() === ANTI_MODE_NAMED_RANGE);
  if (existing) existing.remove();
  ss.setNamedRange(ANTI_MODE_NAMED_RANGE, modeCell);
  return modeCell;
}

function ensureAntiSnapSheet_(sheet, range) {
  const ss = sheet.getParent();
  const snapName = `${ANTI_SNAP_PREFIX}${sheet.getSheetId()}`;
  let snapSheet = ss.getSheetByName(snapName);
  if (!snapSheet) {
    snapSheet = ss.insertSheet(snapName);
    snapSheet.hideSheet();
  }

  const lastRow = range.getLastRow();
  const lastCol = range.getLastColumn();
  runSheetOpWithRetry_(() => ensureSheetSize_(snapSheet, lastRow, lastCol));
  snapSheet.hideSheet();
  return snapSheet;
}

function buildBaseRangeWithExceptions_(sheet, unprotectedRanges) {
  let base = sheet.getDataRange();
  if (!base) return null;

  let minRow = base.getRow();
  let minCol = base.getColumn();
  let maxRow = base.getLastRow();
  let maxCol = base.getLastColumn();

  (unprotectedRanges || []).forEach(r => {
    if (!r || r.getSheet().getSheetId() !== sheet.getSheetId()) return;
    minRow = Math.min(minRow, r.getRow());
    minCol = Math.min(minCol, r.getColumn());
    maxRow = Math.max(maxRow, r.getLastRow());
    maxCol = Math.max(maxCol, r.getLastColumn());
  });

  return sheet.getRange(minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1);
}

function writeSnapshotRange_(snapSheet, range) {
  ensureSheetSize_(snapSheet, range.getLastRow(), range.getLastColumn());
  forEachRangeChunk_(range, ANTI_CHUNK_ROWS_, chunk => {
    const values = chunk.getValues();
    const dst = snapSheet.getRange(chunk.getRow(), chunk.getColumn(), chunk.getNumRows(), chunk.getNumColumns());
    runSheetOpWithRetry_(() => dst.setValues(values));
  });
}

function ensureSheetSize_(sheet, minRows, minCols) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (minRows > maxRows) {
    runSheetOpWithRetry_(() => sheet.insertRowsAfter(maxRows, minRows - maxRows));
  }
  if (minCols > maxCols) {
    runSheetOpWithRetry_(() => sheet.insertColumnsAfter(maxCols, minCols - maxCols));
  }
}

function buildAntiValidationRule_(snapName) {
  const sep = getFormulaSeparator_();
  const formula = `=OR(${ANTI_MODE_NAMED_RANGE}${sep}INDIRECT("'${snapName}'!"&ADDRESS(ROW()${sep}COLUMN()${sep}4))=INDIRECT(ADDRESS(ROW()${sep}COLUMN()${sep}4)))`;
  return SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(false)
    .build();
}

function getFormulaSeparator_() {
  let locale = '';
  try { locale = SpreadsheetApp.getActive().getSpreadsheetLocale() || ''; } catch (_) {}
  if (/^(pt|fr|es|it|de|ru|tr|nl|pl|sv|nb|da|fi|cs|sk|hu|ro|bg|el|uk|lt|lv|et|sl|hr|sr|mk|is|ga|mt)/i.test(locale)) {
    return ';';
  }
  return ',';
}

function applyAntiValidationToRangeFast_(range, rule) {
  forEachRangeChunk_(range, ANTI_CHUNK_ROWS_, chunk => {
    runSheetOpWithRetry_(() => chunk.setDataValidation(rule));
  });
}

function clearAntiValidationInRangeFast_(range) {
  forEachRangeChunk_(range, ANTI_CHUNK_ROWS_, chunk => {
    const dvs = chunk.getDataValidations();
    let changed = false;
    for (let r = 0; r < dvs.length; r++) {
      for (let c = 0; c < dvs[0].length; c++) {
        const dv = dvs[r][c];
        if (isAntiEditValidation_(dv)) {
          dvs[r][c] = null;
          changed = true;
        }
      }
    }
    if (changed) runSheetOpWithRetry_(() => chunk.setDataValidations(dvs));
  });
}

function clearAntiValidationForRanges_(ranges) {
  const list = (ranges || []).filter(Boolean);
  if (!list.length) return;

  const bySheet = new Map();
  list.forEach(r => {
    const sh = r.getSheet();
    const id = sh.getSheetId();
    if (!bySheet.has(id)) bySheet.set(id, { sheet: sh, ranges: [] });
    bySheet.get(id).ranges.push(r);
  });

  bySheet.forEach(entry => {
    const batchSize = 100;
    for (let i = 0; i < entry.ranges.length; i += batchSize) {
      const batch = entry.ranges.slice(i, i + batchSize);
      const notations = batch.map(r => r.getA1Notation());
      try {
        runSheetOpWithRetry_(() => entry.sheet.getRangeList(notations).clearDataValidations());
      } catch (_) {
        batch.forEach(r => {
          runSheetOpWithRetry_(() => r.clearDataValidations());
        });
      }
    }
  });
}

function withAntiEditTemporaryUnblock_(ranges, fn) {
  const list = (ranges || []).filter(Boolean);
  if (!list.length) return fn();

  const ss = list[0].getSheet().getParent();
  const modeCell = ss.getRangeByName(ANTI_MODE_NAMED_RANGE);
  if (!modeCell) return fn();

  clearAntiValidationForRanges_(list);
  const result = fn();
  try { SpreadsheetApp.flush(); } catch (_) {}

  const ruleBySnap = {};
  list.forEach(r => {
    const sheet = r.getSheet();
    const snapSheet = ensureAntiSnapSheet_(sheet, r);
    writeSnapshotRange_(snapSheet, r);
    const snapName = snapSheet.getName();
    const rule = ruleBySnap[snapName] || (ruleBySnap[snapName] = buildAntiValidationRule_(snapName));
    applyAntiValidationToRangeFast_(r, rule);
  });
  return result;
}

function isAntiEditValidation_(dv) {
  try {
    if (!dv) return false;
    if (dv.getCriteriaType && dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) return false;
    const values = dv.getCriteriaValues && dv.getCriteriaValues();
    const formula = values && values[0] ? String(values[0]) : '';
    return formula.indexOf(ANTI_MODE_NAMED_RANGE) !== -1 && formula.indexOf(ANTI_SNAP_PREFIX) !== -1;
  } catch (_) {
    return false;
  }
}

function runSheetOpWithRetry_(fn) {
  const retries = 4;
  for (let i = 0; i < retries; i++) {
    try {
      return fn();
    } catch (err) {
      if (i === retries - 1) throw err;
      Utilities.sleep(250 * (i + 1));
    }
  }
}

function forEachRangeChunk_(range, maxRows, fn) {
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const step = Math.max(1, Number(maxRows) || numRows);
  for (let r = 0; r < numRows; r += step) {
    const rows = Math.min(step, numRows - r);
    const chunk = sheet.getRange(startRow + r, startCol, rows, numCols);
    fn(chunk);
  }
}
