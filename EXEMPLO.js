/** ============================================
 *  ACIONADORES diretos (sem confirmação)
 *  - acionadorVMarket / acionadorComum / efetivarPedido executam direto
 *  - mantém regras visuais (B8) em INTERVALODEINSERÇÃO
 * ============================================ */

/* ========= IDs de planilhas DESTINO ========= */
const TARGET_NR_NAME = 'ACOMPANHAMENTODEPEDIDOS';
const TARGET_SS_IDS = [
  '','', '', '', '', '', '', '', '', '', '',
  '', '', '', '', '', '', '', '', '', ''
];

/* ===================== [ANTI-EDIT] ====================== */
const ANTI_SHEET_NAME = '_ANTI_INDENTIFICACAO_';
const ANTI_MODE_CELL = 'A1';
const ANTI_MODE_NAMED_RANGE = 'ANTI_EDIT_MODE';
const ANTI_BYPASS_CELL = 'A2';
const ANTI_BYPASS_NAMED_RANGE = 'ANTI_EDIT_BYPASS';
const ANTI_SNAP_PREFIX = '_ANTI_SNAP_';

/* ========= Menu ========= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ACIONADORES')
    .addItem('Acionador VMarket', 'acionadorVMarket')
    .addItem('Acionador comum',   'acionadorComum')
    .addItem('Efetivar pedido',   'efetivarPedido')
    .addSeparator()
    .addItem('Liberar permissões', 'liberarPermissoes')
    .addSeparator()
    .addItem('Anti-edit', 'antiEdit_')
    .addToUi();
}

/* ========= onEdit simples (regras visuais do bloco) ========= */
function onEdit(e) {
  try {
    const ss = SpreadsheetApp.getActive();
    const block = ss.getRangeByName('INTERVALODEINSERÇÃO');
    if (!block || !e || !e.range) return;
    const sh = block.getSheet();
    if (e.range.getSheet().getSheetId() !== sh.getSheetId()) return;

    const r0 = block.getRow(), c0 = block.getColumn();
    const r = e.range.getRow(), c = e.range.getColumn();
    if (r < r0 || c < c0 || r > r0 + block.getNumRows() - 1 || c > c0 + block.getNumColumns() - 1) return;

    enforceB8Rule_(block);
  } catch (_) {}
}

/* ========= Wrappers dos três itens de menu ========= */
function acionadorVMarket() { runAcionadorVMarket_(); }
function acionadorComum()   { runAcionadorComum_(); }
function efetivarPedido()   { runEfetivarPedido_(); }

/* ========= ROTINAS REAIS (mesmo código de antes, só renomeado) ========= */

// (antes: acionadorComum)
function runAcionadorComum_() {
  return withAntiEditBypass_(() => {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);
  try {
    const ss = SpreadsheetApp.getActive();
    const srcBlock = ss.getRangeByName('PADRÃOSEMVMARKET');
    const dstBlock = ss.getRangeByName('INTERVALODEINSERÇÃO');

    if (!srcBlock) { safeAlert_('Named range "PADRÃOSEMVMARKET" não encontrado.'); return; }
    if (!dstBlock) { safeAlert_('Named range "INTERVALODEINSERÇÃO" não encontrado.'); return; }

    withAntiEditTemporaryUnblock_([srcBlock, dstBlock], () => {
      const sSh = srcBlock.getSheet();
      const dSh = dstBlock.getSheet();

      const sR = srcBlock.getRow(), sC = srcBlock.getColumn();
      const dR = dstBlock.getRow(), dC = dstBlock.getColumn();

      // Copiar D1 e D5
      dSh.getRange(dR + 0, dC + 3).setValue(sSh.getRange(sR + 0, sC + 3).getValue());
      dSh.getRange(dR + 4, dC + 3).setValue(sSh.getRange(sR + 4, sC + 3).getValue());

      // Copiar B10:E29
      const srcItems = sSh.getRange(sR + 9, sC + 1, 20, 4);
      dSh.getRange(dR + 9, dC + 1, 20, 4).setValues(srcItems.getValues());

      // Limpar origem
      sSh.getRange(sR + 0, sC + 3).clearContent();               // D1
      sSh.getRange(sR + 4, sC + 3).clearContent();               // D5
      sSh.getRange(sR + 9, sC + 1, 20, 4).clearContent();        // B10:E29

      // Carimbo B4
      dSh.getRange(dR + 3, dC + 1).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');

      // F1 = "PE 000001" (sequencial)
      const nextPE = getNextPENumber_();
      dSh.getRange(dR + 0, dC + 5).setValue(nextPE); // F1

      SpreadsheetApp.flush();
      safeToast_('Acionador comum: transferência concluída.', 5);
    });
  } catch (e) {
    safeAlert_('Erro no Acionador comum: ' + (e && e.message ? e.message : e));
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
  });
}

// (antes: acionadorVMarket)
function runAcionadorVMarket_() {
  return withAntiEditBypass_(() => {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getRangeByName('INSERIRMENSAGEMVMARKET');
  const dstBlock = ss.getRangeByName('INTERVALODEINSERÇÃO');

  if (!src) { safeAlert_('Named range "INSERIRMENSAGEMVMARKET" não encontrado.'); return; }
  if (!dstBlock) { safeAlert_('Named range "INTERVALODEINSERÇÃO" não encontrado.'); return; }

  try {
    withAntiEditTemporaryUnblock_([src, dstBlock], () => {
      clearDestination_(dstBlock);

      // Carimbo B4
      const dstSheet = dstBlock.getSheet();
      const topR = dstBlock.getRow();
      const topC = dstBlock.getColumn();
      dstSheet.getRange(topR + 3, topC + 1).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');

      const rawLines = readMessageLinesFromNamed_(src);

  // Remove asteriscos e troca "." por "," apenas fora de números
  const lines = preprocessMessageLines_(rawLines, { 
    stripAsterisks: true,
    replaceDots: 'text-only'   // use 'all' para trocar TODOS os pontos
  });

  if (!lines.length) { ss.toast('Nada para transcrever em INSERIRMENSAGEMVMARKET.', 'ACIONADORES', 5); return; }

  const ctx = parseVMarket_STRICT_noOutros(lines);


      if (ctx.nomeFantasia) dstSheet.getRange(topR + 0, topC + 3).setValue(ctx.nomeFantasia); // D1
      if (ctx.oc)           dstSheet.getRange(topR + 0, topC + 5).setValue(ctx.oc);           // F1
      if (ctx.supplier)     dstSheet.getRange(topR + 4, topC + 3).setValue(ctx.supplier);     // D5

      const START = 10;
      let count = 0;
      ctx.products.forEach((p, i) => {
        const r = START + i;
        if (r > 29) return;

        if (p.quantidade !== undefined && p.quantidade !== '')
          dstSheet.getRange(topR + (r - 1), topC + 1).setValue(p.quantidade); // B
        if (p.gramatura)
          dstSheet.getRange(topR + (r - 1), topC + 2).setValue(p.gramatura);  // C
        if (p.nome)
          dstSheet.getRange(topR + (r - 1), topC + 3).setValue(p.nome);       // D

        if (p.precoUnit !== undefined && p.precoUnit !== '') {
          const cell = dstSheet.getRange(topR + (r - 1), topC + 4); // E
          if (typeof p.precoUnit === 'number') {
            cell.setValue(p.precoUnit);
          } else {
            const n = parseCurrencyToNumber(p.precoUnit);
            if (Number.isFinite(n)) cell.setValue(n); else cell.setValue(p.precoUnit);
          }
        }
        count++;
      });

      ss.toast(`Transcrição concluída: ${count} item(ns).`, 'ACIONADORES', 5);
    });

  } catch (e) {
    safeAlert_('Erro no Acionador VMarket: ' + (e && e.message ? e.message : e));
  } finally {
    try { src.clearContent(); } catch (_) {}
  }
  });
}

// (antes: efetivarPedido)
function runEfetivarPedido_() {
  return withAntiEditBypass_(() => {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);
  try {
    const ss = SpreadsheetApp.getActive();
    const block = ss.getRangeByName('INTERVALODEINSERÇÃO');
    const confNR = ss.getRangeByName('CONFERENCIA');
    if (!block) { safeAlert_('Named range "INTERVALODEINSERÇÃO" não encontrado.'); return; }
    if (!confNR) { safeAlert_('Named range "CONFERENCIA" não encontrado.'); return; }
    const tabNR = ss.getRangeByName('INSERÇÃODADOSDEPEDIDOS');
    const mediaNR = ss.getRangeByName('MEDIADEVALORES');

    const sh = block.getSheet();
    const r0 = block.getRow();
    const c0 = block.getColumn();

      // (0) Validação por TEXTO exibido
      const d1_text  = String(sh.getRange(r0 + 0, c0 + 3).getDisplayValue() || '').trim();   // D1
      const ai1_text = String(confNR.getCell(1, 1).getDisplayValue() || '').trim();          // AI1
      if (d1_text.toLowerCase() !== ai1_text.toLowerCase()) {
        safeToast_('Empresa inválida, script não executado.', 8);
        return;
      }

      const email = String(sh.getRange(r0 + 0, c0 + 1).getDisplayValue() || '').trim();     // B1
      const fileNameBase = String(sh.getRange(r0 + 0, c0 + 5).getDisplayValue() || 'Pedido').trim(); // F1
      const safeName = sanitizeFileName_(fileNameBase) || 'Pedido';

      // (1) PDF + e-mail
      try {
        const pdfBlob = exportRangeAsPdfBlob_(block, safeName);
        if (email && pdfBlob) {
          MailApp.sendEmail({
            to: email,
            subject: `Pedido ${safeName}`,
            htmlBody: `Segue o PDF do pedido <b>${safeName}</b>.`,
            attachments: [pdfBlob.setName(`${safeName}.pdf`)]
          });
        }
      } catch (err) {
        console.error('Falha ao gerar/enviar PDF:', err && err.message ? err.message : err);
      }

      // (2) Copiar bloco para destinos
          // (2) Copiar bloco para destinos (lidos de LIBERAÇÃO IMPORTRANGE!A1)
      try {
        const TARGETS = _getTargetSpreadsheetIds_();
        if (!TARGETS.length) {
          console.error('Nenhum destino encontrado em LIBERAÇÃO IMPORTRANGE!A1 e fallback vazio.');
        }
        TARGETS.forEach(id => {
          if (!id) return;
          try {
            const targetSS = SpreadsheetApp.openById(id);
            const acompNR  = targetSS.getRangeByName(TARGET_NR_NAME);
            if (!acompNR) throw new Error(`Named range "${TARGET_NR_NAME}" não encontrado na planilha destino (${id}).`);
            copyBlockIntoNamedRangeAppend_(block, acompNR);
          } catch (inner) {
            console.error('Falha de cópia p/ destino', id, ':', inner && inner.message ? inner.message : inner);
          }
        });
      } catch (err) {
        console.error('Falha geral na cópia para destinos:', err && err.message ? err.message : err);
      }

      // (3) Inserir dados em INSERÇÃODADOSDEPEDIDOS
      // (3) Inserir/atualizar (idempotente) em INSERÇÃODADOSDEPEDIDOS
      try {
        if (!tabNR) throw new Error('Named range "INSERÇÃODADOSDEPEDIDOS" não encontrado.');

        const tabSh = tabNR.getSheet();
        const tR0 = tabNR.getRow();
        const tC0 = tabNR.getColumn();
        const tRows = tabNR.getNumRows();

        // chaves para detectar duplicado
        const valF1 = String(sh.getRange(r0 + 0, c0 + 5).getDisplayValue() || '').trim(); // A
        const valD1 = String(sh.getRange(r0 + 0, c0 + 3).getDisplayValue() || '').trim(); // B

        // carrega colunas A e B do named range
        const colA = tabSh.getRange(tR0, tC0 + 0, tRows, 1).getDisplayValues().map(r => String(r[0] || '').trim());
        const colB = tabSh.getRange(tR0, tC0 + 1, tRows, 1).getDisplayValues().map(r => String(r[0] || '').trim());

        // procura linha existente com mesma chave (A=F1 e B=D1). Se achar → atualiza; senão → primeira vazia de A
        let insRow = -1;
        for (let i = 0; i < tRows; i++) {
          if (colA[i] !== '' && colA[i] === valF1 && colB[i] === valD1) { insRow = tR0 + i; break; }
        }
        if (insRow === -1) {
          let firstEmpty = colA.findIndex(v => v === '');
          insRow = (firstEmpty === -1) ? (tR0 + tRows) : (tR0 + firstEmpty);
        }

        // garante linhas suficientes
        if (insRow > tabSh.getMaxRows()) tabSh.insertRowsAfter(tabSh.getMaxRows(), insRow - tabSh.getMaxRows());

        // coleta demais valores/fórmulas
        const valB2    = sh.getRange(r0 + 1,  c0 + 1).getDisplayValue();
        const valD5    = sh.getRange(r0 + 4,  c0 + 3).getDisplayValue();
        const valF5    = sh.getRange(r0 + 4,  c0 + 5).getDisplayValue();
        const valB30F30= sh.getRange(r0 + 29, c0 + 1, 1, 5).getDisplayValue(); // B30:F30 (display)
        const valB4    = sh.getRange(r0 + 3,  c0 + 1).getDisplayValue();
        const valB8    = sh.getRange(r0 + 7,  c0 + 1).getDisplayValue();

        const rowRange = tabSh.getRange(insRow, tC0, 1, 23);
        withAntiEditTemporaryUnblock_([rowRange], () => {
          // escreve valores (A..I, H, I) e fórmulas (G, J, K, L, M, N, W)
          tabSh.getRange(insRow, tC0 + 0).setValue(valF1);     // A
          tabSh.getRange(insRow, tC0 + 1).setValue(valD1);     // B
          tabSh.getRange(insRow, tC0 + 2).setValue(valB2);     // C
          tabSh.getRange(insRow, tC0 + 3).setValue(valD5);     // D
          tabSh.getRange(insRow, tC0 + 4).setValue(valF5);     // E
          tabSh.getRange(insRow, tC0 + 5).setValue(valB30F30); // F
          tabSh.getRange(insRow, tC0 + 7).setValue(valB4);     // H
          tabSh.getRange(insRow, tC0 + 8).setValue(valB8);     // I

          const fG = `=IFERROR(VLOOKUP($D${insRow};'DADOS ACUMULADOS'!$E$2:$Q;12;FALSE))`;
          const fJ = `=IF(K${insRow}="Entregue";0;IF(I${insRow}="Variável";"Indeterminado";IFERROR(I${insRow}-TODAY();"")))`;
          const fK = `=IFERROR(VLOOKUP($A${insRow};'DADOS ACUMULADOS'!$AD:$AG;4;FALSE))`;
          const fL = `=IFERROR(VLOOKUP($A${insRow};'DADOS ACUMULADOS'!$AD:$AG;3;FALSE))`;
          const fM = `=IFERROR(VLOOKUP($A${insRow};'DADOS ACUMULADOS'!$AD:$AG;2;FALSE))`;
          const fN = `=T${insRow}-F${insRow}`;
          const fW = `=IFERROR(VLOOKUP($D${insRow};'DADOS ACUMULADOS'!$E$2:$Q;13;FALSE))`;

          tabSh.getRange(insRow, tC0 + 6).setFormula(fG); // G
          tabSh.getRange(insRow, tC0 + 9).setFormula(fJ); // J
          tabSh.getRange(insRow, tC0 + 10).setFormula(fK); // K
          tabSh.getRange(insRow, tC0 + 11).setFormula(fL); // L
          tabSh.getRange(insRow, tC0 + 12).setFormula(fM); // M
          tabSh.getRange(insRow, tC0 + 13).setFormula(fN); // N
          tabSh.getRange(insRow, tC0 + 22).setFormula(fW); // W
        });
      } catch (err) {
        console.error('Falha ao inserir/atualizar em INSERÇÃODADOSDEPEDIDOS:', err && err.message ? err.message : err);
      }

      // (4) MEDIADEVALORES
      try {
        if (!mediaNR) throw new Error('Named range "MEDIADEVALORES" não encontrado.');
    withAntiEditTemporaryUnblock_([mediaNR], () => updateMediaDeValores_(mediaNR, block));
      } catch (err) {
        console.error('Falha ao atualizar MEDIADEVALORES:', err && err.message ? err.message : err);
      }

      // (5) Limpezas finais + fórmula em B8
    withAntiEditTemporaryUnblock_([block], () => {
      try {
        sh.getRange(r0 + 3, c0 + 1).clearContent(); // B4
          sh.getRange(r0 + 0, c0 + 3).clearContent(); // D1
          sh.getRange(r0 + 0, c0 + 5).clearContent(); // F1
          sh.getRange(r0 + 4, c0 + 3).clearContent(); // D5

          const fB8 = `=IFERROR(IF(VLOOKUP($D$5;'DADOS ACUMULADOS'!$E$2:$Q;8;FALSE)="Variável";"Variável";VLOOKUP($D$5;'DADOS ACUMULADOS'!$E$2:$Q;8;FALSE)+$B$4))`;
          const b8 = sh.getRange(r0 + 7, c0 + 1);
          b8.clearContent();
          b8.setFormula(fB8);

          sh.getRange(r0 + 9, c0 + 1, 20, 4).clearContent(); // B10:E29
        } catch (err) {
          console.error('Falha nas limpezas finais:', err && err.message ? err.message : err);
      }
    });

      SpreadsheetApp.flush();
      safeToast_('Efetivar pedido: concluído com sucesso.', 6);
  } catch (e) {
    safeAlert_('Erro no "Efetivar pedido": ' + (e && e.message ? e.message : e));
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
  });
}

/* ========= Helpers reutilizados ========= */

function clearDestination_(dstBlock) {
  const sh = dstBlock.getSheet();
  const r0 = dstBlock.getRow();
  const c0 = dstBlock.getColumn();
  sh.getRange(r0 + 0, c0 + 3).clearContent();         // D1
  sh.getRange(r0 + 0, c0 + 5).clearContent();         // F1
  sh.getRange(r0 + 4, c0 + 3).clearContent();         // D5
  sh.getRange(r0 + 9, c0 + 1, 20, 4).clearContent();  // B10:E29
}
function readMessageLinesFromNamed_(namedRange) {
  const values = namedRange.getDisplayValues();
  if (values.length === 1 && values[0].length === 1) {
    const s = (values[0][0] || '').toString();
    if (s.includes('\n')) return s.split(/\r?\n/).map(x => x.trim());
  }
  const lines = [];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const t = (values[i][j] || '').toString().trim();
      if (t) lines.push(t);
    }
  }
  return lines;
}

function preprocessMessageLines_(lines, opts) {
  opts = opts || {};
  const mode = (opts.replaceDots || 'text-only'); // 'text-only' | 'all' | 'off'
  return (Array.isArray(lines) ? lines : []).map(raw => {
    let t = String(raw || '');
    if (opts.stripAsterisks) t = t.replace(/\*/g, ''); // remove todos os asteriscos
    if (mode === 'all') {
      t = t.replace(/\./g, ',');                        // troca TODOS os pontos por vírgulas
    } else if (mode === 'text-only') {
      t = replaceDotsOutsideNumbers_(t);                // troca só pontos fora de números
    }
    return t.trim();
  });
}

/* Troca "." por "," apenas quando o ponto NÃO está entre dois dígitos.
   Ex.: "12.50" mantém ".", "Item. Novo." vira "Item, Novo," */
function replaceDotsOutsideNumbers_(s) {
  s = String(s || '');
  return s.replace(/\./g, function (dot, i, str) {
    var prev = str[i - 1], next = str[i + 1];
    var prevIsDigit = !!prev && /\d/.test(prev);
    var nextIsDigit = !!next && /\d/.test(next);
    return (prevIsDigit && nextIsDigit) ? '.' : ',';
  });
}

function parseVMarket_STRICT_noOutros(lines) {
  lines = Array.isArray(lines) ? lines : [];
  const ctx = { supplier: '', nomeFantasia: '', oc: '', products: [] };

  ctx.supplier     = extractSupplier(lines);
  ctx.nomeFantasia = findAfter(lines, /^nome\s+fantasia\s*:/i);
  ctx.oc           = findAfter(lines, /^oc\s*:/i);

  const startMarkers = [
    /^seg(?:ue|uem)\s+os\s+produtos/i, /^produtos\s*:/i, /^itens\s*:/i, /^lista\s+de\s+produtos/i
  ];
  let startIdx = lines.findIndex(l => startMarkers.some(r => r.test(String(l || '').trim())));
  if (startIdx < 0) return ctx;

  const endMarkers = [/^quantidade\s+de\s+produtos/i, /^valor\s+total/i, /^\*?\s*confirme/i];

  let current = newItem();
  for (let i = startIdx + 1; i < lines.length; i++) {
    const raw = lines[i];
    const line = String(raw ?? '').trim();
    if (!line) { commit(); continue; }
    if (endMarkers.some(r => r.test(line))) { commit(); break; }

    if (/^gramatura\s*:/i.test(line)) { current.gramatura = afterColon_(line); continue; }
    if (/^(?:preço|preco)\s+unit[áa]rio\s*:/i.test(line)) {
      const v = afterColon_(line);
      const n = parseCurrencyToNumber(v);
      current.precoUnit = Number.isFinite(n) ? n : v;
      continue;
    }
    if (/^quantidade\s*:/i.test(line)) {
  const v = afterColon_(line);
  const q = parseNumberFlexible_(v);
  current.quantidade = Number.isFinite(q) ? q : v;
  continue;
}
    if (/^observa[cç][aã]o(?:es)?\s*:/i.test(line)) { current.observacao = afterColon_(line); continue; }
    if (/^(?:preço|preco)\s+total\s*:/i.test(line)) { /* ignora */ }

    if (line.includes(':')) continue;

    if (hasContent(current)) commit();
    current.nome = removeSufixoOutros_(line);
  }
  commit();
  return ctx;

  function newItem() { return { nome: '', gramatura: '', precoUnit: '', quantidade: '', observacao: '' }; }
  function hasContent(p) { return !!(p.nome || p.gramatura || p.precoUnit !== '' || p.quantidade !== '' || p.observacao); }
  function commit() { if (hasContent(current)) ctx.products.push(current); current = newItem(); }
  function afterColon_(s) { const i = s.indexOf(':'); return i >= 0 ? s.slice(i + 1).trim() : s.trim(); }
}
function removeSufixoOutros_(name) {
  let t = String(name || '').trim();

  // Corta tudo que vier após o ÚLTIMO separador " - " (ou " – " / " — ")
  // Ex.: "CHOPP STELLA ARTOIS BARRIL DE 30L - Outros" → "CHOPP STELLA ARTOIS BARRIL DE 30L"
  //      "CHOPE PATAGONIA IPA 30L - Lote X - Observação" → "CHOPE PATAGONIA IPA 30L"
  const norm = t.replace(/[–—]/g, '-');     // normaliza tipos de traço para '-'
  const idx  = norm.lastIndexOf(' - ');
  if (idx >= 0) t = t.slice(0, idx);

  // Fallback: variações “-Outros” no fim (sem espaço depois do hífen)
  t = t.replace(/\s*[-–—]\s*outros?$/i, '');

  return t.trim();
}


/* ======== Auxiliares visuais/proteção ======== */
function enforceB8Rule_(block) {
  const sh = block.getSheet();
  const r0 = block.getRow();
  const c0 = block.getColumn();
  const b8 = sh.getRange(r0 + 7, c0 + 1);
  const v  = String(b8.getDisplayValue() || '').trim();
  const isVariavel = /vari[áa]vel/i.test(v);

  if (isVariavel) {
    removeAutoB8Protection_(sh, b8);
    b8.setBackground('#b7e1cd'); // verde-claro
  } else {
    ensureAutoB8Protection_(sh, b8);
    b8.setBackground('#ffffff'); // branco
  }
}
function removeAutoB8Protection_(sheet, range) {
  try {
    const prots = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    prots.forEach(p => {
      try {
        if (p.getDescription() === 'AUTO:B8' && p.getRange() && p.getRange().getA1Notation() === range.getA1Notation()) {
          p.remove();
        }
      } catch (_) {}
    });
  } catch (_) {}
}
function ensureAutoB8Protection_(sheet, range) {
  try {
    const prots = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    let p = null;
    for (let i = 0; i < prots.length; i++) {
      try {
        if (prots[i].getDescription() === 'AUTO:B8' && pEquals_(prots[i].getRange(), range)) {
          p = prots[i];
          break;
        }
      } catch (_) {}
    }
    if (!p) {
      p = range.protect();
      p.setDescription('AUTO:B8');
      p.setWarningOnly(false);
      try { p.removeEditors(p.getEditors()); } catch (_) {}
    }
  } catch (_) {}
}
function pEquals_(r1, r2) {
  try { return r1 && r2 && r1.getA1Notation() === r2.getA1Notation() && r1.getSheet().getSheetId() === r2.getSheet().getSheetId(); }
  catch(_) { return false; }
}

/* ======== PDF / cópia visual / append externo ======== */
function exportRangeAsPdfBlob_(range, baseName) {
  const sSh = range.getSheet();
  const ss  = sSh.getParent();
  const tmp = ss.insertSheet(`__TMP_PDF__${Date.now()}`);
  try {
    const numR = range.getNumRows();
    const numC = range.getNumColumns();

    // Copia com visual completo (borda, cores, mesclagens, etc.)
    copyRangeVisual_SameSpreadsheet_(range, tmp, 1, 1);

    // Seleção copiada na aba tmp
    const destRange = tmp.getRange(1, 1, numR, numC);

    // >>> NOVO: normaliza o conteúdo só para o PDF (sem mexer na planilha original)
    _prepareTmpForLabels_(destRange);

    // Ajusta a aba tmp para ter exatamente o tamanho do bloco
    const maxR = tmp.getMaxRows();
    if (maxR > numR) tmp.deleteRows(numR + 1, maxR - numR);
    const maxC = tmp.getMaxColumns();
    if (maxC > numC) tmp.deleteColumns(numC + 1, maxC - numC);

    // Exporta a guia tmp como PDF
    const gid = tmp.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${gid}` +
                `&portrait=false&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenum=UNDEFINED` +
                `&gridlines=false&fzr=false&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25`;

    const token = ScriptApp.getOAuthToken();
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    return res.getBlob().setName(`${baseName}.pdf`);
  } finally {
    try { ss.deleteSheet(tmp); } catch (_) {}
  }
}

function copyRangeVisual_SameSpreadsheet_(srcRange, destSheet, destRow, destCol) {
  const sSh = srcRange.getSheet();
  const sR0 = srcRange.getRow();
  const sC0 = srcRange.getColumn();
  const numR = srcRange.getNumRows();
  const numC = srcRange.getNumColumns();

  const destRange = destSheet.getRange(destRow, destCol, numR, numC);
  for (let j = 0; j < numC; j++) destSheet.setColumnWidth(destCol + j, sSh.getColumnWidth(sC0 + j));
  for (let i = 0; i < numR; i++) destSheet.setRowHeight(destRow + i, sSh.getRowHeight(sR0 + i));
  srcRange.copyTo(destRange, { contentsOnly: false });

  const merges = srcRange.getMergedRanges();
  if (merges && merges.length) {
    merges.forEach(m => {
      const rOff = m.getRow() - sR0;
      const cOff = m.getColumn() - sC0;
      destSheet.getRange(destRow + rOff, destCol + cOff, m.getNumRows(), m.getNumColumns()).merge();
    });
  }
}
function copyBlockIntoNamedRangeAppend_(srcBlock, targetNamedRange) {
  const sSh  = srcBlock.getSheet();
  const tSh  = targetNamedRange.getSheet();
  const tSS  = tSh.getParent();

  return withAntiEditBypassInSpreadsheet_(tSS, () => {
    const sRows = srcBlock.getNumRows();
    const sCols = srcBlock.getNumColumns();

    const sR0 = srcBlock.getRow();
    const sC0 = srcBlock.getColumn();

    const tR0 = targetNamedRange.getRow();
    const tC0 = targetNamedRange.getColumn();

    const lastRowSheet = tSh.getLastRow();
    const scanStart    = tR0;
    const scanEnd      = Math.max(lastRowSheet, tR0);
    const scanHeight   = Math.max(1, scanEnd - scanStart + 1);

    let colA = [];
    try {
      colA = tSh.getRange(scanStart, tC0, scanHeight, 1).getDisplayValues()
                .map(r => String(r[0] || '').trim());
    } catch (e) {
      colA = [];
    }

    let lastNonEmptyIdx = -1;
    for (let i = 0; i < colA.length; i++) if (colA[i] !== '') lastNonEmptyIdx = i;

    const pasteRow = (lastNonEmptyIdx === -1) ? tR0 : (tR0 + lastNonEmptyIdx + 2);

    const needLastRow = pasteRow + sRows - 1;
    if (needLastRow > tSh.getMaxRows()) tSh.insertRowsAfter(tSh.getMaxRows(), needLastRow - tSh.getMaxRows());
    const needLastCol = tC0 + sCols - 1;
    if (needLastCol > tSh.getMaxColumns()) tSh.insertColumnsAfter(tSh.getMaxColumns(), needLastCol - tSh.getMaxColumns());

    const destRange = tSh.getRange(pasteRow, tC0, sRows, sCols);

    try { destRange.breakApart(); } catch (_) {}
    destRange.clear({ contentsOnly: true });

    const values = srcBlock.getValues();
    destRange.setValues(values);

    try { destRange.setNumberFormats(srcBlock.getNumberFormats()); } catch (_) {}
    try { destRange.setBackgrounds(srcBlock.getBackgrounds()); } catch (_) {}
    try { destRange.setFontColors(srcBlock.getFontColors()); } catch (_) {}
    try { destRange.setFontFamilies(srcBlock.getFontFamilies()); } catch (_) {}
    try { destRange.setFontSizes(srcBlock.getFontSizes()); } catch (_) {}
    try { destRange.setFontStyles(srcBlock.getFontStyles()); } catch (_) {}
    try { destRange.setFontWeights(srcBlock.getFontWeights()); } catch (_) {}
    try { destRange.setHorizontalAlignments(srcBlock.getHorizontalAlignments()); } catch (_) {}
    try { destRange.setVerticalAlignments(srcBlock.getVerticalAlignments()); } catch (_) {}
    try { destRange.setWraps(srcBlock.getWraps()); } catch (_) {}

    for (let j = 0; j < sCols; j++) {
      const w = sSh.getColumnWidth(sC0 + j);
      if (w) tSh.setColumnWidth(tC0 + j, w);
    }
    for (let i = 0; i < sRows; i++) {
      const h = sSh.getRowHeight(sR0 + i);
      if (h) tSh.setRowHeight(pasteRow + i, h);
    }

    try {
      const merges = srcBlock.getMergedRanges();
      if (merges && merges.length) {
        for (const m of merges) {
          const rOff = m.getRow()    - sR0;
          const cOff = m.getColumn() - sC0;
          tSh.getRange(pasteRow + rOff, tC0 + cOff, m.getNumRows(), m.getNumColumns()).merge();
        }
      }
    } catch (_) {}
  });
}

/* ======== MEDIADEVALORES ======== */
function updateMediaDeValores_(mediaNR, srcBlock) {
  const mSh = mediaNR.getSheet();
  const mR0 = mediaNR.getRow();
  const mC0 = mediaNR.getColumn();
  const mRows = mediaNR.getNumRows();
  const mCols = mediaNR.getNumColumns();

  const sSh = srcBlock.getSheet();
  const sR0 = srcBlock.getRow();
  const sC0 = srcBlock.getColumn();

  const fornecedor = String(sSh.getRange(sR0 + 4, sC0 + 3).getDisplayValue() || '').trim(); // D5
  if (!fornecedor) return;

  const header = mSh.getRange(mR0 + 0, mC0, 1, mCols).getDisplayValues()[0].map(v => String(v || '').trim());
  let colIdx = header.findIndex(v => v.toLowerCase() === fornecedor.toLowerCase());
  if (colIdx === -1) {
    colIdx = header.findIndex(v => !v);
    if (colIdx === -1 || colIdx + 2 >= mCols) throw new Error('MEDIADEVALORES sem espaço para novo fornecedor no cabeçalho.');
    mSh.getRange(mR0 + 0, mC0 + colIdx).setValue(fornecedor);
  }

  const colProduto = mC0 + colIdx;
  const colGram    = colProduto + 1;
  const colPreco   = colProduto + 2;

  const produtosCol = mSh.getRange(mR0 + 1, colProduto, mRows - 1, 1).getDisplayValues().map(r => String(r[0] || '').trim());

  const maxItems = 20; // B10:E29
  for (let i = 0; i < maxItems; i++) {
    const rowSrc = sR0 + 9 + i;
    const prod = String(sSh.getRange(rowSrc, sC0 + 3).getDisplayValue() || '').trim(); // D
    if (!prod) break;

    const gram = String(sSh.getRange(rowSrc, sC0 + 2).getDisplayValue() || '').trim(); // C
    const precoVal = sSh.getRange(rowSrc, sC0 + 4).getDisplayValue();                  // E
    const precoNum = (typeof precoVal === 'number') ? precoVal : parseCurrencyToNumber(precoVal);

    let foundRel = produtosCol.findIndex(v => v.toLowerCase() === prod.toLowerCase());
    if (foundRel >= 0) {
      const rFound = mR0 + 1 + foundRel;
      const oldVal = mSh.getRange(rFound, colPreco).getDisplayValue();
      const oldNum = (typeof oldVal === 'number') ? oldVal : parseCurrencyToNumber(oldVal);
      const newNum = Number.isFinite(precoNum) ? precoNum : NaN;

      if (Number.isFinite(oldNum) && Number.isFinite(newNum)) {
        mSh.getRange(rFound, colPreco).setValue((oldNum + newNum) / 2);
      } else if (Number.isFinite(newNum)) {
        mSh.getRange(rFound, colPreco).setValue(newNum);
      } else if (precoVal !== '') {
        mSh.getRange(rFound, colPreco).setValue(precoVal);
      }

      const gCell = mSh.getRange(rFound, colGram);
      if (!String(gCell.getDisplayValue() || '').trim() && gram) gCell.setValue(gram);

    } else {
      const firstEmptyRel = produtosCol.findIndex(v => !v);
      if (firstEmptyRel === -1) throw new Error('MEDIADEVALORES sem linhas livres para novos produtos.');
      const rIns = mR0 + 1 + firstEmptyRel;

      mSh.getRange(rIns, colProduto).setValue(prod);
      if (gram) mSh.getRange(rIns, colGram).setValue(gram);
      if (Number.isFinite(precoNum)) mSh.getRange(rIns, colPreco).setValue(precoNum);
      else if (precoVal !== '') mSh.getRange(rIns, colPreco).setValue(precoVal);

      produtosCol[firstEmptyRel] = prod;
    }
  }
}

/* ======== Outras úteis ======== */
function sanitizeFileName_(s) {
  return String(s || '')
    .replace(/[\\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120);
}
function getNextPENumber_() {
  const props = PropertiesService.getDocumentProperties();
  let n = parseInt(props.getProperty('PE_COUNTER') || '0', 10);
  n = isNaN(n) ? 0 : n;
  n += 1;
  props.setProperty('PE_COUNTER', String(n));
  const padded = ('000000' + n).slice(-6);
  return `PE ${padded}`;
}

/* ======== Helpers de UI ======== */
function safeToast_(msg, secs) {
  try { SpreadsheetApp.getActive().toast(msg, 'ACIONADORES', secs || 5); } catch (_) { console.log('[TOAST]', msg); }
}
function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) { console.log('[ALERT]', msg); }
}

/* ===================== [ANTI-EDIT] ===================== */
function antiEditIdentify_() {
  ensureAntiEditSetup_();
  safeToast_('Anti-edit identificado. ANTI_EDIT_MODE pronto.', 5);
}

function antiEdit_() {
  const ss = SpreadsheetApp.getActive();
  deleteAntiEditSheets_(ss);
  antiEditRemove_();
  antiEditIdentify_();
  antiEditApply_();
  safeToast_('Anti-edit concluído.', 5);
}

function antiEditApply_() {
  const ss = SpreadsheetApp.getActive();
  ensureAntiEditSetup_();

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
    writeSnapshot_(snapSheet, baseRange);

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
    writeSnapshot_(snapSheet, range);

    const snapName = snapSheet.getName();
    const rule = ruleBySnap[snapName] || (ruleBySnap[snapName] = buildAntiValidationRule_(snapName));
    applyAntiValidationToRangeFast_(range, rule);
  });

  safeToast_('Anti-edit aplicado.', 5);
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

  safeToast_('Anti-edit removido.', 5);
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
  const ss = SpreadsheetApp.getActive();
  return ensureAntiEditSetupForSpreadsheet_(ss);
}

function ensureAntiEditSetupForSpreadsheet_(ss) {
  let sheet = ss.getSheetByName(ANTI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANTI_SHEET_NAME);
    sheet.hideSheet();
  }

  const modeCell = sheet.getRange(ANTI_MODE_CELL);
  if (modeCell.getValue() === '') modeCell.setValue(false);
  const bypassCell = sheet.getRange(ANTI_BYPASS_CELL);
  if (bypassCell.getValue() === '') bypassCell.setValue(false);
  sheet.hideSheet();

  const existing = ss.getNamedRanges().find(nr => nr.getName() === ANTI_MODE_NAMED_RANGE);
  if (existing) existing.remove();
  ss.setNamedRange(ANTI_MODE_NAMED_RANGE, modeCell);
  const existingBypass = ss.getNamedRanges().find(nr => nr.getName() === ANTI_BYPASS_NAMED_RANGE);
  if (existingBypass) existingBypass.remove();
  ss.setNamedRange(ANTI_BYPASS_NAMED_RANGE, bypassCell);
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

function writeSnapshot_(snapSheet, range) {
  const values = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  ensureSheetSize_(snapSheet, range.getLastRow(), range.getLastColumn());
  runSheetOpWithRetry_(() => {
    snapSheet.getRange(startRow, startCol, range.getNumRows(), range.getNumColumns()).setValues(values);
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
  const formula = `=OR(${ANTI_MODE_NAMED_RANGE}${sep}${ANTI_BYPASS_NAMED_RANGE}${sep}INDIRECT("'${snapName}'!"&ADDRESS(ROW()${sep}COLUMN()${sep}4))=INDIRECT(ADDRESS(ROW()${sep}COLUMN()${sep}4)))`;
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
    runSheetOpWithRetry_(() => chunk.clearDataValidations());
  });
}

function clearAntiValidationForRanges_(ranges) {
  const list = (ranges || []).filter(Boolean);
  if (!list.length) return;
  const sheet = list[0].getSheet();
  const batchSize = 100;
  for (let i = 0; i < list.length; i += batchSize) {
    const batch = list.slice(i, i + batchSize);
    const batchNotations = batch.map(r => r.getA1Notation());
    try {
      runSheetOpWithRetry_(() => {
        const rangeList = sheet.getRangeList(batchNotations);
        rangeList.clearDataValidations();
      });
    } catch (_) {
      batch.forEach(r => {
        runSheetOpWithRetry_(() => r.clearDataValidations());
      });
    }
  }
}

function withAntiEditTemporaryUnblock_(ranges, fn, opts) {
  const list = (ranges || []).filter(Boolean);
  if (!list.length) return fn();
  const ss = list[0].getSheet().getParent();
  return withAntiEditBypassInSpreadsheet_(ss, () => {
    list.forEach(r => {
      try { r.clearDataValidations(); } catch (_) {}
    });

    const result = fn();
    try { SpreadsheetApp.flush(); } catch (_) {}
    const ruleBySnap = {};
    list.forEach(r => {
      const sheet = r.getSheet();
      const snapSheet = ensureAntiSnapSheet_(sheet, r);
      writeSnapshot_(snapSheet, r);
      const snapName = snapSheet.getName();
      const rule = ruleBySnap[snapName] || (ruleBySnap[snapName] = buildAntiValidationRule_(snapName));
      applyAntiValidationToRangeFast_(r, rule);
    });
    if (opts && opts.reapplyAll) {
      antiEditApply_();
    }
    return result;
  });
}

function withAntiEditBypass_(fn) {
  return withAntiEditBypassInSpreadsheet_(SpreadsheetApp.getActive(), fn);
}

function withAntiEditBypassInSpreadsheet_(spreadsheet, fn) {
  if (!spreadsheet) return fn();
  if (isAntiBypassEnabledInSpreadsheet_(spreadsheet)) return fn();
  let ok = false;
  try {
    setAntiBypassModeInSpreadsheet_(spreadsheet, true);
    ok = true;
    return fn();
  } finally {
    try {
      if (ok) refreshAntiSnapshotsForSpreadsheet_(spreadsheet);
      if (ok) setAntiBypassModeInSpreadsheet_(spreadsheet, false);
    } catch (_) {}
  }
}

function setAntiBypassMode_(value) {
  setAntiBypassModeInSpreadsheet_(SpreadsheetApp.getActive(), value);
}

function setAntiBypassModeInSpreadsheet_(spreadsheet, value) {
  ensureAntiEditSetupForSpreadsheet_(spreadsheet);
  const named = spreadsheet.getRangeByName(ANTI_BYPASS_NAMED_RANGE);
  if (named) named.setValue(!!value);
}

function isAntiBypassEnabled_() {
  return isAntiBypassEnabledInSpreadsheet_(SpreadsheetApp.getActive());
}

function isAntiBypassEnabledInSpreadsheet_(spreadsheet) {
  try {
    const named = spreadsheet.getRangeByName(ANTI_BYPASS_NAMED_RANGE);
    return named ? !!named.getValue() : false;
  } catch (_) {
    return false;
  }
}

function refreshAntiSnapshots_() {
  refreshAntiSnapshotsForSpreadsheet_(SpreadsheetApp.getActive());
}

function refreshAntiSnapshotsForSpreadsheet_(spreadsheet) {
  const sheetProtections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
  const rangeProtections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];

  sheetProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const sheet = p.getRange().getSheet();
    const unprotected = p.getUnprotectedRanges ? (p.getUnprotectedRanges() || []) : [];
    const baseRange = buildBaseRangeWithExceptions_(sheet, unprotected);
    if (!baseRange) return;
    const snapSheet = ensureAntiSnapSheet_(sheet, baseRange);
    writeSnapshot_(snapSheet, baseRange);
  });

  rangeProtections.forEach(p => {
    if (p.isWarningOnly && p.isWarningOnly()) return;
    const range = p.getRange();
    if (!range) return;
    const snapSheet = ensureAntiSnapSheet_(range.getSheet(), range);
    writeSnapshot_(snapSheet, range);
  });
}

function runSheetOpWithRetry_(fn) {
  const retries = 3;
  for (let i = 0; i < retries; i++) {
    try {
      return fn();
    } catch (err) {
      if (i === retries - 1) throw err;
      Utilities.sleep(300 * (i + 1));
    }
  }
}

const ANTI_CHUNK_ROWS_ = 200;

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

/* ======== Suporte ======== */
function parseCurrencyToNumber(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/\s+/g, '').replace(/[R$\u00A0]/gi, '');
  return parseNumberFlexible_(s);
}

// Converte números escritos com ponto OU vírgula como decimal.
// Aceita formatos: 15.1, 15,1, 1.234,56, 1,234.56, R$ 32,98, etc.
function parseNumberFlexible_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').trim();

  // remove moeda/espacos
  s = s.replace(/[R$\u00A0\s]/gi, '');

  const hasComma = s.includes(',');
  const hasDot   = s.includes('.');

  if (hasComma && hasDot) {
    // Se os dois existem, decide pelo ÚLTIMO separador como decimal
    const lastComma = s.lastIndexOf(',');
    const lastDot   = s.lastIndexOf('.');
    if (lastComma > lastDot) {
      // vírgula é decimal → remove pontos (milhar) e troca vírgula por ponto
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      // ponto é decimal → remove vírgulas (milhar)
      s = s.replace(/,/g, '');
    }
  } else if (hasComma) {
    // só vírgula → vírgula é decimal
    s = s.replace(',', '.');
  } else {
    // só ponto (ou nenhum). Se tiver >1 ponto, trate como milhares → remove todos
    if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
    // com 1 ponto só, assume ponto como decimal (deixa como está)
  }

  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}

function findAfter(lines, regex) {
  for (const raw of (lines || [])) {
    const line = String(raw || '').trim();
    if (regex.test(line)) {
      const i = line.indexOf(':');
      return i >= 0 ? line.slice(i + 1).trim() : line;
    }
  }
  return '';
}

function extractSupplier(lines) {
  // 1) Caminho "rotulado"
  const labeled =
    findAfter(lines, /^fornecedor\s*:/i) ||
    findAfter(lines, /^supplier\s*:/i)   ||
    findAfter(lines, /^empresa\s*:/i);
  if (labeled) return labeled;

  // 2) Fallback: frase do tipo "… com a/o/os/as NOME DO FORNECEDOR"
  // Agora aceita parênteses, aspas e sinais comuns no nome.
  const HEAD_SCAN = 15; // varre só o cabeçalho
  for (const raw of (lines || []).slice(0, HEAD_SCAN)) {
    const s = String(raw || '').trim();

    // Ex.: "gostaria de realizar o pedido com a (PIMENTA PENONI) PICANCIA DE MINAS LTDA"
    // Captura TUDO até o fim da linha (permitindo (), aspas “ ” ‘ ’, hífens etc.)
    const m = s.match(
      /\bcom\s+(?:o|a|os|as)?\s*([A-Za-zÀ-ÿ0-9 .,&\/\-ºª()"'“”’]+?)\s*$/i
    );

    if (m && m[1]) {
      let sup = m[1].trim();
      // remove pontuação final solta (ex.: vírgulas/pontos/traços ao fim)
      sup = sup.replace(/\s*[.,;:–—-]\s*$/,'').trim();
      return sup;
    }
  }

  return '';
}

/** ====== [LIBERAÇÃO IMPORTRANGE!A1 → URL/ID do Slides] ====== **/

// Lê a URL de LIBERAÇÃO IMPORTRANGE!A1 (rich text, HYPERLINK ou texto cru)
function _readUrlFromLiberacaoA1_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('LIBERAÇÃO IMPORTRANGE') ||
             ss.getSheets().find(s => String(s.getName() || '').toLowerCase() === 'liberação importrange');
  if (!sh) return '';
  const cell = sh.getRange(1, 1); // A1

  // 1) RichText (link aplicado no texto – total ou parcial)
  try {
    const rt = cell.getRichTextValue();
    if (rt) {
      const whole = rt.getLinkUrl && rt.getLinkUrl();
      if (whole) return whole;
      const runs = rt.getRuns && rt.getRuns();
      if (runs && runs.length) {
        for (var i = 0; i < runs.length; i++) {
          var u = runs[i].getLinkUrl && runs[i].getLinkUrl();
          if (u) return u;
        }
      }
    }
  } catch (_) {}

  // 2) Fórmula HYPERLINK("url","texto") ou HYPERLINK("url";"texto")
  try {
    const f = String(cell.getFormula() || '');
    const m = f.match(/HYPERLINK\(\s*"([^"]+)"\s*[,;]\s*"/i);
    if (m && m[1]) return m[1];
  } catch (_) {}

  // 3) Valor exibido começando com http(s)
  const disp = String(cell.getDisplayValue() || '').trim();
  if (/^https?:\/\//i.test(disp)) return disp;

  return '';
}

// Extrai ID de arquivos do Google Drive a partir de várias formas de URL
function _extractDriveFileIdFromUrl_(url) {
  url = String(url || '').trim();
  if (!url) return '';
  const tries = [
    /\/d\/([a-zA-Z0-9_-]{20,})/,          // .../d/<ID>/...
    /[?&]id=([a-zA-Z0-9_-]{20,})/,        // ...?id=<ID>
    /\/folders\/([a-zA-Z0-9_-]{20,})/,    // .../folders/<ID> (se algum dia usar pasta)
  ];
  for (const re of tries) {
    const m = url.match(re);
    if (m && m[1]) return m[1];
  }
  return '';
}

// Retorna os IDs de destino a partir do hyperlink em A1; se não achar, cai no fallback TARGET_SS_IDS
function _getTargetSpreadsheetIds_() {
  const url = _readUrlFromLiberacaoA1_();
  const id  = _extractDriveFileIdFromUrl_(url);
  if (id) return [id];
  // fallback para o array fixo, caso A1 não esteja configurado
  return (Array.isArray(TARGET_SS_IDS) ? TARGET_SS_IDS : []).filter(Boolean);
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

function _prepareTmpForLabels_(destRange) {
  try {
    // 1) Troca FORNECEDOR: -> END:
    destRange.createTextFinder('FORNECEDOR:')
      .matchCase(true)
      .useRegularExpression(false)
      .replaceAllWith('END:');
  } catch (_) {}

  try {
    // 2) Remove quebras de linha dentro das células
    destRange.createTextFinder('[\\r\\n]+')
      .useRegularExpression(true)
      .replaceAllWith(' ');
  } catch (_) {}

  try {
    // 3) Comprimi espaços múltiplos, só estética
    destRange.createTextFinder(' {2,}')
      .useRegularExpression(true)
      .replaceAllWith(' ');
  } catch (_) {}

  try {
    // 4) Sem quebra de linha e cortando excesso
    if (destRange.setWrapStrategy) {
      destRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    } else {
      // fallback para versões antigas
      destRange.setWrap(false);
    }
  } catch (_) {}
}