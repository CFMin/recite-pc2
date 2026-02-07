const STORAGE_KEY = 'recite_v1';

function uid() {
  return Math.random().toString(16).slice(2) + Date.now().toString(16);
}

function isSelectionInside(elm) {
  try {
    const sel = window.getSelection?.();
    if (!sel || sel.rangeCount === 0) return false;
    const r = sel.getRangeAt(0);
    if (!r || r.collapsed) return false;
    const c = r.commonAncestorContainer;
    const node = c && c.nodeType === Node.ELEMENT_NODE ? c : c?.parentElement;
    return !!node && elm.contains(node);
  } catch {
    return false;
  }
}

function applyAnswerHighlight(color) {
  const qa = getCurrentQa();
  if (!qa) return;
  if (!ui.viewAnswer) return;

  // Only support direct rich-html view (not mask/check rendering)
  const richRoot = ui.viewAnswer.querySelector('.rich-answer-editable');
  if (!richRoot) {
    alert('当前答案不是富文本展示状态，无法进行框选标注。请停止播放/关闭遮罩或清空检测后再试。');
    return;
  }

  if (!isSelectionInside(richRoot)) {
    alert('请先在答案区域框选一段文字，再点击标注颜色。');
    return;
  }

  const sel = window.getSelection();
  const range = sel.getRangeAt(0);
  if (!range || range.collapsed) return;

  try {
    const mk = document.createElement('mark');
    mk.setAttribute('data-color', String(color || 'yellow'));
    const frag = range.extractContents();
    mk.appendChild(frag);
    range.insertNode(mk);

    // Merge nested marks created by repeated operations
    richRoot.querySelectorAll('mark mark').forEach((inner) => {
      const outer = inner.parentElement;
      if (!outer) return;
      while (inner.firstChild) outer.insertBefore(inner.firstChild, inner);
      inner.remove();
    });

    // Persist
    const nextHtml = sanitizeAnswerHtml(richRoot.innerHTML || '');
    if (nextHtml.includes('▇▇▇▇▇▇')) {
      alert('检测到答案处于遮罩状态（包含占位符）。请先关闭遮罩/清空检测后再进行标注。');
      return;
    }
    const nextText = htmlToPlainTextPreserveLines(nextHtml);
    upsertQa({ ...qa, answerHtml: nextHtml, answerText: nextText, updatedAt: nowIso() });
    saveState();
    render();
    updateMatches();
    try { sel.removeAllRanges(); } catch {}
  } catch {
    alert('该段文字跨越了复杂结构，无法直接标注。请尽量在同一段落内框选。');
  }
}

function clearAnswerHighlights() {
  const qa = getCurrentQa();
  if (!qa) return;
  if (!ui.viewAnswer) return;

  const richRoot = ui.viewAnswer.querySelector('.rich-answer-editable');
  if (!richRoot) {
    alert('当前答案不是富文本展示状态，无法清除标注。请停止播放/关闭遮罩或清空检测后再试。');
    return;
  }

  // If selection is inside, only clear marks intersecting selection; otherwise clear all.
  const onlySel = isSelectionInside(richRoot);
  if (!onlySel) {
    richRoot.querySelectorAll('mark').forEach((m) => {
      const p = m.parentNode;
      if (!p) return;
      while (m.firstChild) p.insertBefore(m.firstChild, m);
      m.remove();
    });
  } else {
    const sel = window.getSelection();
    const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    if (!range) return;
    const marks = Array.from(richRoot.querySelectorAll('mark'));
    marks.forEach((m) => {
      try {
        if (!range.intersectsNode(m)) return;
      } catch {
        // Older browsers may throw; fall back to clearing all when selection exists
      }
      const p = m.parentNode;
      if (!p) return;
      while (m.firstChild) p.insertBefore(m.firstChild, m);
      m.remove();
    });
  }

  const nextHtml = sanitizeAnswerHtml(richRoot.innerHTML || '');
  if (nextHtml.includes('▇▇▇▇▇▇')) {
    alert('检测到答案处于遮罩状态（包含占位符）。请先关闭遮罩/清空检测后再清除标注。');
    return;
  }
  const nextText = htmlToPlainTextPreserveLines(nextHtml);
  upsertQa({ ...qa, answerHtml: nextHtml, answerText: nextText, updatedAt: nowIso() });
  saveState();
  render();
  updateMatches();
}

function repairMaskPlaceholderForCurrentQa() {
  const qa = getCurrentQa();
  if (!qa) return;
  const htmlHas = String(qa.answerHtml || '').includes('▇▇▇▇▇▇');
  const textHas = String(qa.answerText || '').includes('▇▇▇▇▇▇');
  if (!htmlHas && !textHas) {
    alert('未检测到遮罩占位符污染，无需修复。');
    return;
  }

  const ok = confirm('检测到答案已被“遮罩占位符”污染（▇▇▇▇▇▇）。修复会移除这些占位符；若占位符覆盖了原文，原文需要你手动补回。是否继续？');
  if (!ok) return;

  let nextHtml = '';
  let nextText = '';

  if (!textHas && String(qa.answerText || '').trim()) {
    nextText = String(qa.answerText || '');
    nextHtml = sanitizeAnswerHtml(escapeHtml(nextText).replace(/\n/g, '<br>'));
  } else {
    const safeHtml = sanitizeAnswerHtml(normalizeImportedHtml(qa.answerHtml || ''));
    nextHtml = safeHtml.split('▇▇▇▇▇▇').join('');
    nextText = htmlToPlainTextPreserveLines(nextHtml).replace(/▇▇▇▇▇▇/g, '').trim();
    nextHtml = sanitizeAnswerHtml(nextHtml);
  }

  upsertQa({ ...qa, answerHtml: nextHtml, answerText: nextText, updatedAt: nowIso() });
  saveState();
  ui.inputRecited.value = '';
  resetReciteCheck();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function getReciteBucketId(level) {
  const lv = Number(level);
  if (lv === 0) return RECITE_BUCKETS[0].id;
  if (lv === 1) return RECITE_BUCKETS[1].id;
  if (lv === 2) return RECITE_BUCKETS[2].id;
  return RECITE_BUCKETS[1].id;
}

function markCurrentQaToReciteLevel(level) {
  const qa = getCurrentQa();
  if (!qa) return;
  if (!listState.collectionId) {
    alert('请先进入一个合集，再进行归档。');
    return;
  }

  flushFocusPoints();
  commitActiveQaTime('markLevel');
  stopPlayer();

  const sourceCollectionId = listState.collectionId;
  const base = state.qas.filter((x) => (x.collectionId || DEFAULT_COLLECTION_ID) === sourceCollectionId);
  const idx = base.findIndex((x) => x.id === qa.id);
  const nextQaId = idx >= 0 && idx + 1 < base.length ? base[idx + 1].id : idx > 0 ? base[idx - 1].id : null;

  const counts = qa.reciteLevelCounts && typeof qa.reciteLevelCounts === 'object' ? { ...qa.reciteLevelCounts } : { 0: 0, 1: 0, 2: 0 };
  const k = String(Number(level));
  counts[k] = (Number(counts[k]) || 0) + 1;

  const targetCollectionId = getReciteBucketId(level);
  upsertQa({
    ...qa,
    reciteLevel: Number(level),
    reciteLevelCounts: counts,
    collectionId: targetCollectionId,
    updatedAt: nowIso(),
  });

  // Move on within the source collection
  if (nextQaId) {
    setCurrentQa(nextQaId);
    fillEditorFromCurrent();
  } else {
    state.progress.currentQaId = null;
    saveState();
    render();
    fillEditorFromCurrent();
    updateMatches();
  }
}

function deleteCurrentCollectionQas() {
  const colId = listState.collectionId;
  if (!colId) return;
  const col = (state.collections || []).find((c) => c.id === colId);
  const name = col ? col.name : '当前合集';
  const count = state.qas.filter((q) => (q.collectionId || DEFAULT_COLLECTION_ID) === colId).length;
  if (!count) {
    alert('当前合集没有 QA。');
    return;
  }
  const ok = confirm(`确认删除合集“${name}”下的全部 ${count} 条 QA？此操作不可恢复！`);
  if (!ok) return;

  flushFocusPoints();
  commitActiveQaTime('deleteCollectionAll');
  stopPlayer();

  const toDelete = new Set(state.qas.filter((q) => (q.collectionId || DEFAULT_COLLECTION_ID) === colId).map((q) => q.id));
  state.qas = state.qas.filter((q) => !toDelete.has(q.id));
  listState.selected.clear();
  listState.page = 1;

  if (state.progress.currentQaId && toDelete.has(state.progress.currentQaId)) {
    state.progress.currentQaId = null;
  }

  saveState();
  resetReciteCheck();
  resetCardCheck();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function deleteCollectionById(colId) {
  if (!colId) return;
  if (PROTECTED_COLLECTION_IDS.has(colId)) {
    alert('该合集不能删除。');
    return;
  }

  const col = (state.collections || []).find((c) => c.id === colId);
  const name = col ? col.name : '该合集';
  const count = state.qas.filter((q) => (q.collectionId || DEFAULT_COLLECTION_ID) === colId).length;
  const ok = confirm(`确认删除合集“${name}”？将同时删除其中全部 ${count} 条 QA，且不可恢复！`);
  if (!ok) return;

  flushFocusPoints();
  commitActiveQaTime('deleteCollection');
  stopPlayer();

  // Delete QAs in this collection
  const toDelete = new Set(state.qas.filter((q) => (q.collectionId || DEFAULT_COLLECTION_ID) === colId).map((q) => q.id));
  state.qas = state.qas.filter((q) => !toDelete.has(q.id));

  // Delete the collection itself
  state.collections = (state.collections || []).filter((c) => c.id !== colId);
  ensureCollections();

  // Fix current selection if needed
  if (state.progress.currentQaId && toDelete.has(state.progress.currentQaId)) {
    state.progress.currentQaId = null;
  }
  if (state.progress.currentCollectionId === colId) {
    state.progress.currentCollectionId = null;
  }
  if (listState.collectionId === colId) {
    listState.collectionId = null;
  }

  listState.selected.clear();
  listState.page = 1;
  saveState();
  resetReciteCheck();
  resetCardCheck();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function deleteAllQas() {
  if (!state.qas.length) return;
  const ok = confirm('确认删除全部 QA？此操作不可恢复。');
  if (!ok) return;
  commitActiveQaTime('deleteAll');
  stopPlayer();
  state.qas = [];
  state.collections = [{ id: DEFAULT_COLLECTION_ID, name: '默认合集', createdAt: nowIso() }];
  state.progress.currentQaId = null;
  state.progress.currentCollectionId = null;
  listState.collectionId = null;
  listState.selected.clear();
  listState.query = '';
  ui.inputQaSearch.value = '';
  listState.page = 1;
  ui.inputQuestion.value = '';
  ui.inputAnswer.value = '';
  ui.inputRecited.value = '';
  resetReciteCheck();
  resetCardCheck();
  saveState();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function buildPlayerStepsForQa(qa) {
  const mainSteps = buildStepsForQa(qa, { review: false });
  
  // Insert full reading before groups
  const beforeFullReadSteps = [];
  const beforeCount = clamp(Number(state.settings.fullReadBeforeGroups) || 1, 0, 5);
  for (let i = 0; i < beforeCount; i++) {
    beforeFullReadSteps.push({
      qaId: qa.id,
      kind: 'speak',
      text: qa.question || '',
      isQuestion: true,
      review: false,
      isFullRead: true,
      fullReadType: 'before',
    });
    beforeFullReadSteps.push({
      qaId: qa.id,
      kind: 'speak',
      text: qa.answerText || '',
      isQuestion: false,
      review: false,
      isFullRead: true,
      fullReadType: 'before',
    });
  }
  
  // Insert a step to read the question first (for compatibility)
  const questionStep = {
    qaId: qa.id,
    kind: 'speak',
    text: qa.question || '',
    isQuestion: true,
    review: false,
  };
  
  let steps = [...beforeFullReadSteps, questionStep, ...mainSteps];
  
  // Insert full reading after groups
  const afterFullReadSteps = [];
  const afterCount = clamp(Number(state.settings.fullReadAfterGroups) || 1, 0, 5);
  for (let i = 0; i < afterCount; i++) {
    afterFullReadSteps.push({
      qaId: qa.id,
      kind: 'speak',
      text: qa.question || '',
      isQuestion: true,
      review: false,
      isFullRead: true,
      fullReadType: 'after',
    });
    afterFullReadSteps.push({
      qaId: qa.id,
      kind: 'speak',
      text: qa.answerText || '',
      isQuestion: false,
      review: false,
      isFullRead: true,
      fullReadType: 'after',
    });
  }
  steps = [...steps, ...afterFullReadSteps];

  const prevId = prevQaId(qa.id);
  if (state.settings.reviewPrevAfterEach && prevId) {
    const prevQa = getQaById(prevId);
    if (prevQa) {
      const reviewSteps = buildStepsForQa(prevQa, { review: true });
      const reviewQuestionStep = {
        qaId: prevQa.id,
        kind: 'speak',
        text: prevQa.question || '',
        isQuestion: true,
        review: true,
      };
      const reviewRepeatCount = clamp(Number(state.settings.reviewPrevRepeatCount) || 1, 1, 5);
      const reviewRounds = [];
      for (const st of reviewSteps) {
        if (st.round < reviewRepeatCount) reviewRounds.push(st);
      }
      steps = [...steps, reviewQuestionStep, ...reviewRounds];
    }
  }
  return steps;
}

function splitSegments(sentence) {
  const t = (sentence || '').trim();
  if (!t) return [];
  const out = [];
  let buf = '';
  for (const ch of t) {
    buf += ch;
    if (/[，,;；、]/.test(ch)) {
      const p = buf.trim();
      if (p) out.push(p);
      buf = '';
    }
  }
  if (buf.trim()) out.push(buf.trim());

  const res = [];
  for (const p of out) {
    if (p.length > 24) {
      const mid = Math.ceil(p.length / 2);
      res.push(p.slice(0, mid), p.slice(mid));
    } else {
      res.push(p);
    }
  }
  return res.map((x) => x.trim()).filter(Boolean);
}

function buildSegments(sentences) {
  const segments = [];
  const bySentence = [];
  for (let si = 0; si < sentences.length; si++) {
    const segs = splitSegments(sentences[si]);
    bySentence[si] = [];
    for (let sj = 0; sj < segs.length; sj++) {
      const globalIndex = segments.length;
      const item = { globalIndex, sentenceIndex: si, segmentIndex: sj, text: segs[sj] };
      segments.push(item);
      bySentence[si].push(item);
    }
    if (!segs.length) {
      const globalIndex = segments.length;
      const item = { globalIndex, sentenceIndex: si, segmentIndex: 0, text: (sentences[si] || '').trim() };
      segments.push(item);
      bySentence[si].push(item);
    }
  }
  return { segments, bySentence };
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function nowIso() {
  return new Date().toISOString();
}

function defaultState() {
  return {
    collections: [
      {
        id: 'default',
        name: '默认合集',
        createdAt: nowIso(),
      },
    ],
    qas: [
      {
        id: uid(),
        question: '示例：什么是强化学习？',
        answerText: '强化学习是一类通过与环境交互来学习策略的方法。智能体在状态下选择动作并获得奖励。目标是最大化长期累积回报。',
        collectionId: 'default',
        createdAt: nowIso(),
        updatedAt: nowIso(),
      },
    ],
    settings: {
      groupSize: 3,
      repeatPerGroup: 4,
      reviewPrevAfterEach: true,
      sentenceDelimiters: '。！？!?',
      ttsEnabled: true,
      ttsVoiceUri: '',
      autoPlayNextQa: false,
      forceReciteCheck: false,
      reciteOnlyHighlights: false,
      qaReciteSideBySide: false,
      rate: 1.0,
      volume: 1.0,
      threshold: 0.65,
      useOllama: false,
      ollamaUrl: 'http://localhost:11434',
      ollamaModel: 'qwen2.5:3b',
      fullReadBeforeGroups: 1,
      fullReadAfterGroups: 1,
      reviewPrevRepeatCount: 1,
    },
    progress: {
      currentQaId: null,
      currentCollectionId: null,
    },
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return defaultState();
    if (!Array.isArray(parsed.qas)) return defaultState();
    parsed.settings = { ...defaultState().settings, ...(parsed.settings || {}) };
    parsed.progress = { ...defaultState().progress, ...(parsed.progress || {}) };
    if (!Array.isArray(parsed.collections)) parsed.collections = defaultState().collections;
    return parsed;
  } catch {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function download(filename, content) {
  const blob = new Blob([content], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(s) {
  return (s ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function maskPlaceholder(text) {
  return '▇▇▇▇▇▇';
}

function normalizeText(s) {
  return (s || '')
    .toLowerCase()
    .replace(/[\s\u3000]+/g, '')
    .replace(/[，。！？、；：,.!?;:"'“”‘’（）()【】\[\]{}<>《》\-—_]/g, '');
}

function bigrams(s) {
  const t = normalizeText(s);
  if (!t) return [];
  if (t.length === 1) return [t];
  const res = [];
  for (let i = 0; i < t.length - 1; i++) res.push(t.slice(i, i + 2));
  return res;
}

function diceSimilarity(a, b) {
  const A = bigrams(a);
  const B = bigrams(b);
  if (!A.length || !B.length) return 0;
  const m = new Map();
  for (const x of A) m.set(x, (m.get(x) || 0) + 1);
  let overlap = 0;
  for (const y of B) {
    const c = m.get(y) || 0;
    if (c > 0) {
      overlap++;
      m.set(y, c - 1);
    }
  }
  return (2 * overlap) / (A.length + B.length);
}

function splitSentences(text) {
  const raw = (text || '').replace(/\r\n/g, '\n');

  const delimsRaw = String(state?.settings?.sentenceDelimiters || '。！？!?').replace(/\s+/g, '');
  const delims = delimsRaw ? [...new Set(delimsRaw.split(''))].join('') : '。！？!?';
  const cls = delims.replace(/[\\\]\[\-\^]/g, (m) => `\\${m}`);

  let marked = raw;
  try {
    marked = raw.replace(new RegExp(`([${cls}])`, 'g'), '$1\n');
  } catch {
    marked = raw;
  }

  const parts = marked
    .split(/\n+/g)
    .map((s) => s.trim())
    .filter(Boolean);
  return parts.length ? parts : raw.trim() ? [raw.trim()] : [];
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function el(id) {
  return document.getElementById(id);
}

function isTypingTarget(node) {
  const el = node;
  if (!el) return false;
  const tag = (el.tagName || '').toUpperCase();
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
  if (el.isContentEditable) return true;
  return false;
}

function normalizeImportedHtml(html) {
  const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
  const body = doc.body;
  if (!body) return '';

  const cssToMarkColor = (raw) => {
    const v = String(raw || '').trim().toLowerCase();
    if (!v) return 'yellow';
    if (v.includes('yellow')) return 'yellow';
    if (v.includes('green')) return 'green';
    if (v.includes('cyan') || v.includes('aqua') || v.includes('turquoise')) return 'cyan';
    if (v.includes('magenta') || v.includes('fuchsia') || v.includes('purple') || v.includes('violet')) return 'magenta';

    const hex = v.match(/#([0-9a-f]{3}|[0-9a-f]{6})/i)?.[0];
    if (hex) {
      let h = hex.slice(1);
      if (h.length === 3) h = h.split('').map((c) => c + c).join('');
      const r = parseInt(h.slice(0, 2), 16);
      const g = parseInt(h.slice(2, 4), 16);
      const b = parseInt(h.slice(4, 6), 16);
      if (r > 220 && g > 220 && b < 120) return 'yellow';
      if (g > 180 && r < 140 && b < 140) return 'green';
      if (b > 180 && g > 180 && r < 140) return 'cyan';
      if (r > 160 && b > 160 && g < 150) return 'magenta';
      return 'yellow';
    }

    const rgb = v.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
    if (rgb) {
      const r = Number(rgb[1]);
      const g = Number(rgb[2]);
      const b = Number(rgb[3]);
      if (r > 220 && g > 220 && b < 120) return 'yellow';
      if (g > 180 && r < 140 && b < 140) return 'green';
      if (b > 180 && g > 180 && r < 140) return 'cyan';
      if (r > 160 && b > 160 && g < 150) return 'magenta';
      return 'yellow';
    }

    return 'yellow';
  };

  const parseStyle = (style) => {
    const s = String(style || '').toLowerCase();
    const bg = s.match(/background(?:-color)?\s*:\s*([^;]+)/i)?.[1]?.trim();
    const fw = s.match(/font-weight\s*:\s*([^;]+)/i)?.[1]?.trim();
    const bold = fw === 'bold' || Number(fw) >= 600;
    return { bg, bold };
  };

  // Convert mammoth spans with background/font-weight into our markup
  body.querySelectorAll('span').forEach((sp) => {
    const { bg, bold } = parseStyle(sp.getAttribute('style') || '');
    if (!bg && !bold) return;

    let root = null;
    let leaf = null;

    if (bg) {
      const mk = doc.createElement('mark');
      mk.setAttribute('data-color', cssToMarkColor(bg));
      root = mk;
      leaf = mk;
    }

    if (bold) {
      const st = doc.createElement('strong');
      if (!root) {
        root = st;
        leaf = st;
      } else {
        leaf.appendChild(st);
        leaf = st;
      }
    }

    leaf.innerHTML = sp.innerHTML;
    sp.replaceWith(root);
  });

  // Strip scripts/styles
  body.querySelectorAll('script,style').forEach((n) => n.remove());
  return body.innerHTML;
}

function sanitizeAnswerHtml(html) {
  const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
  const body = doc.body;
  if (!body) return '';

  const allowed = new Set(['STRONG', 'B', 'MARK', 'BR', 'DIV', 'P']);
  const walker = doc.createTreeWalker(body, NodeFilter.SHOW_ELEMENT, null);
  const toProcess = [];
  while (walker.nextNode()) toProcess.push(walker.currentNode);

  for (const node of toProcess) {
    const tag = (node.tagName || '').toUpperCase();
    if (!allowed.has(tag)) {
      const frag = doc.createDocumentFragment();
      while (node.firstChild) frag.appendChild(node.firstChild);
      node.replaceWith(frag);
      continue;
    }
    // Remove all attributes (keep clean), but preserve mark color
    Array.from(node.attributes || []).forEach((a) => {
      const name = String(a.name || '').toLowerCase();
      if (tag === 'MARK' && name === 'data-color') return;
      node.removeAttribute(a.name);
    });
  }

  return body.innerHTML;
}

function htmlToPlainTextPreserveLines(html) {
  const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
  const body = doc.body;
  if (!body) return '';
  // Insert newlines for block-ish separators
  body.querySelectorAll('br').forEach((br) => br.replaceWith(doc.createTextNode('\n')));
  body.querySelectorAll('p,div,h1,h2,h3,h4,h5,h6,li').forEach((n) => {
    n.appendChild(doc.createTextNode('\n'));
  });
  return (body.textContent || '')
    .replace(/[\u00a0]/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n')
    .trim();
}

function extractAnswerHtmlFromRichEditor() {
  if (!ui.inputAnswerRich) return '';
  return sanitizeAnswerHtml(ui.inputAnswerRich.innerHTML || '');
}

function extractHighlightSegmentsFromAnswerHtml(answerHtml) {
  try {
    const rawHtml = sanitizeAnswerHtml(normalizeImportedHtml(answerHtml || ''));
    const doc = new DOMParser().parseFromString(rawHtml || '', 'text/html');
    const marks = Array.from(doc.body?.querySelectorAll('mark') || []);
    const segs = [];
    for (const m of marks) {
      const t = String(m.textContent || '').replace(/[\u00a0]/g, ' ').trim();
      if (!t) continue;
      segs.push(t);
    }
    return segs;
  } catch {
    return [];
  }
}

function getReciteSegmentsForQa(qa) {
  const onlyHi = !!state?.settings?.reciteOnlyHighlights;
  if (onlyHi && qa?.answerHtml) {
    const parts = extractHighlightSegmentsFromAnswerHtml(qa.answerHtml);
    const segments = parts.map((t, i) => ({ globalIndex: i, sentenceIndex: 0, segmentIndex: i, text: t }));
    const bySentence = [segments];
    return { segments, bySentence };
  }
  const sentences = splitSentences(qa?.answerText || '');
  return buildSegments(sentences);
}

function setRichEditorHtml(html) {
  if (!ui.inputAnswerRich) return;
  const cleaned = sanitizeAnswerHtml(normalizeImportedHtml(html));
  ui.inputAnswerRich.innerHTML = cleaned;
}

function getPlainAnswerFromRichEditor() {
  return htmlToPlainTextPreserveLines(extractAnswerHtmlFromRichEditor());
}

const ui = {
  qaList: el('qaList'),
  inputQaSearch: el('inputQaSearch'),
  checkSelectAll: el('checkSelectAll'),
  btnDeleteCollectionAll: el('btnDeleteCollectionAll'),
  btnDeleteSelected: el('btnDeleteSelected'),
  btnClearSelection: el('btnClearSelection'),
  btnDeleteAll: el('btnDeleteAll'),
  btnCollectionBack: el('btnCollectionBack'),
  collectionTitle: el('collectionTitle'),
  btnPagePrev: el('btnPagePrev'),
  pageInfo: el('pageInfo'),
  btnPageNext: el('btnPageNext'),
  inputQuestion: el('inputQuestion'),
  inputAnswer: el('inputAnswer'),
  inputAnswerRich: el('inputAnswerRich'),
  btnAnswerBold: el('btnAnswerBold'),
  btnAnswerHighlight: el('btnAnswerHighlight'),
  btnAnswerClearFormat: el('btnAnswerClearFormat'),
  btnSave: el('btnSave'),
  btnNew: el('btnNew'),
  btnDelete: el('btnDelete'),

  inputGroupSize: el('inputGroupSize'),
  inputRepeat: el('inputRepeat'),
  inputSentenceDelims: el('inputSentenceDelims'),
  checkReviewPrev: el('checkReviewPrev'),
  checkTts: el('checkTts'),
  selectTtsVoice: el('selectTtsVoice'),
  checkAutoPlayNext: el('checkAutoPlayNext'),
  checkForceRecite: el('checkForceRecite'),
  checkQaReciteSideBySide: el('checkQaReciteSideBySide'),
  inputFullReadBefore: el('inputFullReadBefore'),
  inputFullReadAfter: el('inputFullReadAfter'),
  inputReviewPrevRepeat: el('inputReviewPrevRepeat'),
  inputRate: el('inputRate'),
  inputVolume: el('inputVolume'),
  inputThreshold: el('inputThreshold'),
  btnApplySettings: el('btnApplySettings'),

  currentTitle: el('currentTitle'),
  statusLine: el('statusLine'),
  viewQuestion: el('viewQuestion'),
  viewAnswer: el('viewAnswer'),
  inputFocusPoints: el('inputFocusPoints'),
  qaReciteWrap: el('qaReciteWrap'),

  btnCardMode: el('btnCardMode'),
  btnCardPrev: el('btnCardPrev'),
  btnCardFlip: el('btnCardFlip'),
  btnCardNext: el('btnCardNext'),
  btnMaskToggleAll: el('btnMaskToggleAll'),

  btnAnsMarkYellow: el('btnAnsMarkYellow'),
  btnAnsMarkGreen: el('btnAnsMarkGreen'),
  btnAnsMarkCyan: el('btnAnsMarkCyan'),
  btnAnsMarkMagenta: el('btnAnsMarkMagenta'),
  btnAnsMarkClear: el('btnAnsMarkClear'),

  ansMarkPalette: el('ansMarkPalette'),

  btnStart: el('btnStart'),
  btnPause: el('btnPause'),
  btnStop: el('btnStop'),
  btnTtsTest: el('btnTtsTest'),
  btnPrev: el('btnPrev'),
  btnNext: el('btnNext'),

  btnMarkLevel0: el('btnMarkLevel0'),
  btnMarkLevel1: el('btnMarkLevel1'),
  btnMarkLevel2: el('btnMarkLevel2'),
  reciteLevelStats: el('reciteLevelStats'),

  btnRecStart: el('btnRecStart'),
  btnRecStop: el('btnRecStop'),
  btnCheck: el('btnCheck'),
  btnRepairMaskPlaceholder: el('btnRepairMaskPlaceholder'),
  checkReciteOnlyHighlights: el('checkReciteOnlyHighlights'),
  inputRecited: el('inputRecited'),
  matchSummary: el('matchSummary'),
  speechHint: el('speechHint'),

  inputAsk: el('inputAsk'),
  btnAsk: el('btnAsk'),
  askAnswer: el('askAnswer'),
  checkUseOllama: el('checkUseOllama'),
  inputOllamaUrl: el('inputOllamaUrl'),
  inputOllamaModel: el('inputOllamaModel'),

  btnExport: el('btnExport'),
  btnExportDocx: el('btnExportDocx'),
  fileImport: el('fileImport'),

  qaTimer: el('qaTimer'),
  currentQaTimer: el('currentQaTimer'),
  btnQaTimerPause: el('btnQaTimerPause'),
};

function hideAnsMarkPalette() {
  if (!ui.ansMarkPalette) return;
  ui.ansMarkPalette.style.display = 'none';
}

function showAnsMarkPaletteNearSelection() {
  if (!ui.ansMarkPalette) return;
  if (!ui.viewAnswer) return;

  const qa = getCurrentQa();
  if (!qa) return hideAnsMarkPalette();

  const richRoot = ui.viewAnswer.querySelector('.rich-answer-editable');
  if (!richRoot) return hideAnsMarkPalette();

  const sel = window.getSelection?.();
  if (!sel || sel.rangeCount === 0) return hideAnsMarkPalette();
  const range = sel.getRangeAt(0);
  if (!range || range.collapsed) return hideAnsMarkPalette();

  if (!isSelectionInside(richRoot)) return hideAnsMarkPalette();

  const rect = range.getBoundingClientRect();
  if (!rect || (!rect.width && !rect.height)) return hideAnsMarkPalette();

  const prev = ui.ansMarkPalette.style.display;
  ui.ansMarkPalette.style.display = 'flex';
  const pw = ui.ansMarkPalette.offsetWidth || 190;
  const ph = ui.ansMarkPalette.offsetHeight || 40;
  ui.ansMarkPalette.style.display = prev || 'none';

  const pad = 8;
  let top = rect.top - ph - pad;
  if (top < 8) top = rect.bottom + pad;

  let left = rect.right - pw;
  if (left < 8) left = rect.left;
  if (left + pw > window.innerWidth - 8) left = window.innerWidth - pw - 8;

  ui.ansMarkPalette.style.left = `${Math.round(left)}px`;
  ui.ansMarkPalette.style.top = `${Math.round(top)}px`;
  ui.ansMarkPalette.style.display = 'flex';
}

let state = loadState();
if (!state.progress.currentQaId && state.qas.length) state.progress.currentQaId = state.qas[0].id;

const DEFAULT_COLLECTION_ID = 'default';

const RECITE_BUCKETS = [
  { id: 'recite_level_0', name: '0-不会' },
  { id: 'recite_level_1', name: '1-一般' },
  { id: 'recite_level_2', name: '2-会了' },
];

const PROTECTED_COLLECTION_IDS = new Set([DEFAULT_COLLECTION_ID, ...RECITE_BUCKETS.map((x) => x.id)]);

function ensureCollections() {
  if (!Array.isArray(state.collections)) {
    state.collections = [{ id: DEFAULT_COLLECTION_ID, name: '默认合集', createdAt: nowIso() }];
  }
  if (!state.collections.find((c) => c.id === DEFAULT_COLLECTION_ID)) {
    state.collections.unshift({ id: DEFAULT_COLLECTION_ID, name: '默认合集', createdAt: nowIso() });
  }

  // Ensure recite bucket collections exist
  for (const b of RECITE_BUCKETS) {
    if (!state.collections.find((c) => c.id === b.id)) {
      state.collections.push({ id: b.id, name: b.name, createdAt: nowIso() });
    }
  }

  state.qas = (state.qas || []).map((qa) => {
    if (!qa || typeof qa !== 'object') return qa;
    if (!qa.collectionId) return { ...qa, collectionId: DEFAULT_COLLECTION_ID };
    return qa;
  });
  const known = new Set(state.collections.map((c) => c.id));
  state.qas.forEach((qa) => {
    if (qa?.collectionId && !known.has(qa.collectionId)) {
      state.collections.push({ id: qa.collectionId, name: '未命名合集', createdAt: nowIso() });
      known.add(qa.collectionId);
    }
  });
}

ensureCollections();
saveState();

let _focusPointsSaveTimer = null;
let _focusPointsDirtyQaId = null;

function flushFocusPoints() {
  if (!ui.inputFocusPoints) return;
  const qa = getCurrentQa();
  if (!qa) {
    _focusPointsDirtyQaId = null;
    return;
  }
  if (_focusPointsDirtyQaId !== qa.id) return;
  const v = String(ui.inputFocusPoints.value || '');
  if ((qa.focusPoints || '') === v) {
    _focusPointsDirtyQaId = null;
    return;
  }
  upsertQa({ ...qa, focusPoints: v, updatedAt: nowIso() });
  _focusPointsDirtyQaId = null;
}

function scheduleSaveFocusPoints() {
  const qa = getCurrentQa();
  if (!qa) return;
  _focusPointsDirtyQaId = qa.id;
  if (_focusPointsSaveTimer) clearTimeout(_focusPointsSaveTimer);
  _focusPointsSaveTimer = setTimeout(() => {
    _focusPointsSaveTimer = null;
    flushFocusPoints();
  }, 400);
}

const qaTimerState = {
  activeQaId: null,
  sessionStartMs: null,
  paused: false,
  pausedAtMs: null,
  pausedAccumMs: 0,
  intervalId: null,
};

function ensureQaTimeMap() {
  if (!state.progress) state.progress = { currentQaId: null };
  if (!state.progress.qaTimes) state.progress.qaTimes = {};
  if (!state.progress.qaTimeHistory) state.progress.qaTimeHistory = {};
}

function formatDurationMs(ms) {
  const s = Math.max(0, Math.floor(ms / 1000));
  const m = Math.floor(s / 60);
  const ss = String(s % 60).padStart(2, '0');
  return `${m}:${ss}`;
}

function getQaTotalMs(qaId) {
  ensureQaTimeMap();
  return Number(state.progress.qaTimes?.[qaId] || 0);
}

function getQaHistoryList(qaId) {
  ensureQaTimeMap();
  const list = state.progress.qaTimeHistory?.[qaId];
  return Array.isArray(list) ? list : [];
}

function pushQaHistory(qaId, ms) {
  ensureQaTimeMap();
  if (!qaId) return;
  const v = Math.max(0, Math.floor(Number(ms) || 0));
  if (!v) return;
  const prev = getQaHistoryList(qaId);
  const next = [v, ...prev].slice(0, 10);
  state.progress.qaTimeHistory[qaId] = next;
}

function computeCurrentSessionMs() {
  if (!qaTimerState.activeQaId) return 0;
  if (qaTimerState.paused) return qaTimerState.pausedAccumMs;
  if (!qaTimerState.sessionStartMs) return qaTimerState.pausedAccumMs;
  return qaTimerState.pausedAccumMs + Math.max(0, Date.now() - qaTimerState.sessionStartMs);
}

function commitActiveQaTime(reason) {
  ensureQaTimeMap();
  if (!qaTimerState.activeQaId) return;
  const delta = computeCurrentSessionMs();
  if (delta > 0) {
    const prev = getQaTotalMs(qaTimerState.activeQaId);
    state.progress.qaTimes[qaTimerState.activeQaId] = prev + delta;
    pushQaHistory(qaTimerState.activeQaId, delta);
    saveState();
  }
  // Reset session (next QA or next time we enter current QA)
  qaTimerState.sessionStartMs = null;
  qaTimerState.paused = false;
  qaTimerState.pausedAtMs = null;
  qaTimerState.pausedAccumMs = 0;
}

function resetQaTimerSession() {
  qaTimerState.sessionStartMs = qaTimerState.activeQaId ? Date.now() : null;
  qaTimerState.paused = false;
  qaTimerState.pausedAtMs = null;
  qaTimerState.pausedAccumMs = 0;
}

function computeActiveQaDisplayMs() {
  // Only show current session time, not accumulated total
  return computeCurrentSessionMs();
}

function renderQaTimer() {
  if (!ui.qaTimer || !ui.currentQaTimer) return;
  const qa = getCurrentQa();
  const ms = qa ? computeActiveQaDisplayMs() : 0;
  const text = formatDurationMs(ms);
  ui.qaTimer.textContent = text;
  ui.currentQaTimer.textContent = `本题计时：${text}`;

  if (ui.btnQaTimerPause) {
    ui.btnQaTimerPause.textContent = qaTimerState.paused ? '计时继续' : '计时暂停';
    ui.btnQaTimerPause.disabled = !qa;
  }

  const warn = ms >= 5 * 60 * 1000;
  ui.qaTimer.classList.toggle('qa-timer-warn', warn);
  ui.currentQaTimer.classList.toggle('qa-timer-warn', warn);
}

let tickAudioCtx = null;
function playTick() {
  try {
    if (!tickAudioCtx) tickAudioCtx = new (window.AudioContext || window.webkitAudioContext)();
    const ctx = tickAudioCtx;
    if (ctx.state === 'suspended') ctx.resume?.();
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.type = 'sine';
    o.frequency.value = 880;
    // Lower volume to reduce interference with TTS
    g.gain.value = 0.0001;
    o.connect(g);
    g.connect(ctx.destination);
    const t0 = ctx.currentTime;
    // Softer envelope to avoid cutting off TTS
    g.gain.exponentialRampToValueAtTime(0.02, t0 + 0.01);
    g.gain.exponentialRampToValueAtTime(0.0001, t0 + 0.1);
    o.start(t0);
    o.stop(t0 + 0.12);
  } catch {}
}

function startQaTimerForCurrent() {
  ensureQaTimeMap();
  const qa = getCurrentQa();
  const qaId = qa?.id || null;

  if (qaTimerState.activeQaId !== qaId) {
    commitActiveQaTime('switch');
    qaTimerState.activeQaId = qaId;
    resetQaTimerSession(); // Start fresh at 0 for new QA
  } else {
    if (qaId && !qaTimerState.sessionStartMs && !qaTimerState.paused) qaTimerState.sessionStartMs = Date.now();
  }

  if (!qaTimerState.intervalId) {
    qaTimerState.intervalId = setInterval(() => {
      renderQaTimer();
      // Allow tick sound during TTS but at lower volume to reduce interference
      const shouldTick = !!qaTimerState.activeQaId && !qaTimerState.paused && !document.hidden;
      if (shouldTick) playTick();
    }, 1000);
  }
  renderQaTimer();
}

function toggleQaTimerPause() {
  const qa = getCurrentQa();
  if (!qa) return;
  if (!qaTimerState.activeQaId) startQaTimerForCurrent();

  if (qaTimerState.paused) {
    qaTimerState.paused = false;
    qaTimerState.sessionStartMs = Date.now();
  } else {
    qaTimerState.pausedAccumMs = computeCurrentSessionMs();
    qaTimerState.sessionStartMs = null;
    qaTimerState.paused = true;
  }
  renderQaTimer();
}

function setCurrentQa(id) {
  flushFocusPoints();
  commitActiveQaTime();
  state.progress.currentQaId = id;
  ui.inputRecited.value = '';
  resetReciteCheck();
  resetCardCheck();
  saveState();
  render();
  updateMatches();
  startQaTimerForCurrent(); // This will reset timer to 0 for new QA
}

try {
  window.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      // Leaving the page shouldn't count as switching QA, but we also shouldn't keep accumulating.
      // So we pause accumulation by committing the current session to pausedAccum.
      if (qaTimerState.activeQaId && !qaTimerState.paused && qaTimerState.sessionStartMs) {
        qaTimerState.pausedAccumMs = computeCurrentSessionMs();
        qaTimerState.sessionStartMs = null;
      }
    } else {
      if (qaTimerState.activeQaId && !qaTimerState.paused) {
        qaTimerState.sessionStartMs = Date.now();
      }
      renderQaTimer();
    }
  });
  window.addEventListener('beforeunload', () => {
    flushFocusPoints();
    commitActiveQaTime('unload');
  });
} catch {}

ui.inputFocusPoints?.addEventListener('input', () => scheduleSaveFocusPoints());
ui.inputFocusPoints?.addEventListener('blur', () => flushFocusPoints());

// Rich editor wiring (bold/highlight)
if (ui.inputAnswerRich) {
  const syncPlain = () => {
    ui.inputAnswer.value = getPlainAnswerFromRichEditor();
  };

  ui.inputAnswerRich.addEventListener('input', () => syncPlain());

  const exec = (cmd, val) => {
    try {
      ui.inputAnswerRich.focus();
      document.execCommand(cmd, false, val);
      syncPlain();
    } catch {}
  };

  ui.btnAnswerBold?.addEventListener('click', () => exec('bold'));
  ui.btnAnswerHighlight?.addEventListener('click', () => {
    // Most browsers accept 'hiliteColor'; fallback to 'backColor'
    exec('hiliteColor', '#ffcc66');
    exec('backColor', '#ffcc66');
  });
  ui.btnAnswerClearFormat?.addEventListener('click', () => {
    exec('removeFormat');
    // unwrap <mark> if any remains
    try {
      ui.inputAnswerRich.querySelectorAll('mark').forEach((m) => {
        const frag = document.createDocumentFragment();
        while (m.firstChild) frag.appendChild(m.firstChild);
        m.replaceWith(frag);
      });
    } catch {}
    syncPlain();
  });
}

let _ttsVoicesSig = '';
function listVoices() {
  try {
    return ('speechSynthesis' in window && speechSynthesis.getVoices) ? speechSynthesis.getVoices() || [] : [];
  } catch {
    return [];
  }
}

function populateTtsVoiceSelect() {
  if (!ui.selectTtsVoice) return;
  const voices = listVoices();
  const sig = voices.map((v) => v.voiceURI).join('|');
  if (sig === _ttsVoicesSig && ui.selectTtsVoice.options.length) {
    const desired = String(state.settings.ttsVoiceUri || '');
    if (ui.selectTtsVoice.value !== desired) ui.selectTtsVoice.value = desired;
    return;
  }
  _ttsVoicesSig = sig;

  const desired = String(state.settings.ttsVoiceUri || '');
  ui.selectTtsVoice.innerHTML = '';

  const optAuto = document.createElement('option');
  optAuto.value = '';
  optAuto.textContent = '自动（中文优先）';
  ui.selectTtsVoice.appendChild(optAuto);

  voices.forEach((v) => {
    const opt = document.createElement('option');
    opt.value = v.voiceURI || '';
    opt.textContent = `${v.name || 'voice'} (${v.lang || ''})`;
    ui.selectTtsVoice.appendChild(opt);
  });

  ui.selectTtsVoice.value = desired;
}

try {
  if ('speechSynthesis' in window) {
    speechSynthesis.onvoiceschanged = () => {
      populateTtsVoiceSelect();
    };
  }
} catch {}

let player = {
  running: false,
  paused: false,
  qaId: null,
  mainQaId: null,
  steps: [],
  stepIndex: 0,
  activeSentenceGlobalIndex: null,
  speechUtterance: null,
  reviewMode: false,
};

let cardCheck = {
  enabled: false,
  flipped: false,
  index: 0,
};

let reciteCheck = {
  lockedSegments: new Set(),
  pointerSegment: 0,
  lastUtterance: '',
  maskMode: false,
  manualAuto: false,
};

let maskState = {
  showAll: false,
};

let listState = {
  pageSize: 5,
  page: 1,
  query: '',
  selected: new Set(),
  collectionId: state.progress?.currentCollectionId || null,
};

function getQaById(id) {
  return state.qas.find((q) => q.id === id) || null;
}

function getCurrentQa() {
  return getQaById(state.progress.currentQaId);
}


function getFilteredQas() {
  const q = (listState.query || '').trim().toLowerCase();
  const base = listState.collectionId
    ? state.qas.filter((x) => x.collectionId === listState.collectionId)
    : state.qas;
  if (!q) return base;
  return base.filter((x) => {
    const a = (x.answerText || '').toLowerCase();
    const b = (x.question || '').toLowerCase();
    return a.includes(q) || b.includes(q);
  });
}

function getCollectionsFiltered() {
  const q = (listState.query || '').trim().toLowerCase();
  const cols = Array.isArray(state.collections) ? state.collections : [];
  if (!q) return cols;
  return cols.filter((c) => String(c.name || '').toLowerCase().includes(q));
}

function setActiveCollection(id) {
  listState.collectionId = id;
  listState.page = 1;
  listState.selected.clear();
  ui.checkSelectAll.checked = false;
  ui.checkSelectAll.indeterminate = false;
  if (!state.progress) state.progress = { currentQaId: null, currentCollectionId: null };
  state.progress.currentCollectionId = id;
  saveState();
  render();
}

ui.btnCollectionBack?.addEventListener('click', () => {
  listState.query = '';
  ui.inputQaSearch.value = '';
  setActiveCollection(null);
});

function normalizeImportedCollectionName(filename) {
  const base = String(filename || '').trim() || '导入';
  return base.replace(/\.(json|docx)$/i, '');
}

function createCollection(name) {
  ensureCollections();
  const col = {
    id: uid(),
    name: String(name || '未命名合集'),
    createdAt: nowIso(),
  };
  state.collections.unshift(col);
  return col;
}

function getPagedQas() {
  const filtered = getFilteredQas();
  const totalPages = Math.max(1, Math.ceil(filtered.length / listState.pageSize));
  listState.page = clamp(listState.page, 1, totalPages);
  const start = (listState.page - 1) * listState.pageSize;
  const items = filtered.slice(start, start + listState.pageSize);
  return { filtered, items, totalPages };
}

function resetReciteCheck() {
  reciteCheck.lockedSegments = new Set();
  reciteCheck.pointerSegment = 0;
  reciteCheck.lastUtterance = '';
  reciteCheck.maskMode = false;
  reciteCheck.manualAuto = false;
  maskState.showAll = false;
  updateManualCheckButton();
}

function updateManualCheckButton() {
  if (!ui.btnCheck) return;
  ui.btnCheck.textContent = reciteCheck.manualAuto ? '自动输入检测：关' : '用输入检测';
}

function resetCardCheck() {
  cardCheck.flipped = false;
  cardCheck.index = 0;
}

function renderQaList() {
  const cur = state.progress.currentQaId;

  // Collection list mode
  if (!listState.collectionId) {
    const cols = getCollectionsFiltered();
    const counts = new Map();
    state.qas.forEach((qa) => {
      const cid = qa.collectionId || DEFAULT_COLLECTION_ID;
      counts.set(cid, (counts.get(cid) || 0) + 1);
    });

    if (ui.btnCollectionBack) ui.btnCollectionBack.disabled = true;
    if (ui.collectionTitle) ui.collectionTitle.textContent = '合集';
    if (ui.pageInfo) ui.pageInfo.textContent = `共 ${cols.length} 个合集`;
    ui.btnPagePrev.disabled = true;
    ui.btnPageNext.disabled = true;
    ui.checkSelectAll.disabled = true;
    if (ui.btnDeleteCollectionAll) ui.btnDeleteCollectionAll.disabled = true;
    ui.btnDeleteSelected.disabled = true;
    ui.btnClearSelection.disabled = true;

    ui.qaList.innerHTML = cols
      .map((c) => {
        const n = counts.get(c.id) || 0;
        const disableDel = PROTECTED_COLLECTION_IDS.has(c.id) ? 'disabled' : '';
        return `
          <div class="qa-item" data-col-id="${c.id}">
            <div class="qa-item-row">
              <div class="qa-item-main">
                <div class="q">${escapeHtml(c.name || '（未命名合集）')}</div>
                <div class="a">${n} 条</div>
              </div>
              <button class="btn btn-danger btn-mini" data-col-del-id="${c.id}" ${disableDel}>删除</button>
            </div>
          </div>
        `;
      })
      .join('');

    ui.qaList.querySelectorAll('[data-col-id]').forEach((node) => {
      node.addEventListener('click', () => {
        const id = node.getAttribute('data-col-id');
        if (!id) return;
        setActiveCollection(id);
        const first = state.qas.find((q) => q.collectionId === id);
        if (first) {
          state.progress.currentQaId = first.id;
          saveState();
          fillEditorFromCurrent();
        }
        render();
        updateMatches();
      });
    });

    ui.qaList.querySelectorAll('[data-col-del-id]').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const id = btn.getAttribute('data-col-del-id');
        if (!id) return;
        deleteCollectionById(id);
      });
    });
    return;
  }

  // QA list mode (inside a collection)
  ui.checkSelectAll.disabled = false;
  if (ui.btnDeleteCollectionAll) ui.btnDeleteCollectionAll.disabled = false;
  if (ui.btnCollectionBack) ui.btnCollectionBack.disabled = false;
  const col = state.collections.find((c) => c.id === listState.collectionId);
  if (ui.collectionTitle) ui.collectionTitle.textContent = col ? col.name : '合集';

  const { items, filtered, totalPages } = getPagedQas();

  ui.pageInfo.textContent = `第 ${listState.page}/${totalPages} 页（共 ${filtered.length} 条）`;
  ui.btnPagePrev.disabled = listState.page <= 1;
  ui.btnPageNext.disabled = listState.page >= totalPages;

  ui.qaList.innerHTML = items
    .map((qa) => {
      const active = qa.id === cur ? 'active' : '';
      const checked = listState.selected.has(qa.id) ? 'checked' : '';
      const hist = getQaHistoryList(qa.id);
      const histText = hist.length ? hist.map((ms) => formatDurationMs(ms)).join(' / ') : '';
      return `
        <div class="qa-item ${active}" data-id="${qa.id}">
          <div class="qa-item-row">
            <input class="qa-check" type="checkbox" data-check-id="${qa.id}" ${checked} />
            <div class="qa-item-main">
              <div class="q">${escapeHtml(qa.question || '（无问题）')}</div>
              <div class="a">${escapeHtml((qa.answerText || '').slice(0, 90))}</div>
              ${histText ? `<div class="qa-time-hist">近10次：${escapeHtml(histText)}</div>` : ''}
            </div>
          </div>
        </div>
      `;
    })
    .join('');

  ui.qaList.querySelectorAll('.qa-item').forEach((node) => {
    node.addEventListener('click', (e) => {
      const t = e.target;
      if (t && t.matches && t.matches('input[type="checkbox"]')) return;
      const id = node.getAttribute('data-id');
      if (!id) return;

      if ((listState.query || '').trim()) {
        listState.query = '';
        ui.inputQaSearch.value = '';
        const idxInAll = state.qas.findIndex((x) => x.id === id);
        if (idxInAll >= 0) listState.page = Math.floor(idxInAll / listState.pageSize) + 1;
      }

      setCurrentQa(id);
      fillEditorFromCurrent();
    });
  });

  ui.qaList.querySelectorAll('input[data-check-id]').forEach((node) => {
    node.addEventListener('click', (e) => e.stopPropagation());
    node.addEventListener('change', () => {
      const id = node.getAttribute('data-check-id');
      if (!id) return;
      if (node.checked) listState.selected.add(id);
      else listState.selected.delete(id);
      updateSelectAllState();
    });
  });

  updateSelectAllState();
}

function updateSelectAllState() {
  const { items } = getPagedQas();
  const ids = items.map((x) => x.id);
  const selectedCount = ids.filter((id) => listState.selected.has(id)).length;
  const total = ids.length;
  ui.checkSelectAll.indeterminate = selectedCount > 0 && selectedCount < total;
  ui.checkSelectAll.checked = total > 0 && selectedCount === total;
  ui.btnDeleteSelected.disabled = listState.selected.size === 0;
  ui.btnClearSelection.disabled = listState.selected.size === 0;
}

function renderSettings() {
  const s = state.settings;
  ui.inputGroupSize.value = String(s.groupSize);
  ui.inputRepeat.value = String(s.repeatPerGroup);
  ui.inputSentenceDelims.value = String(s.sentenceDelimiters || '');
  ui.checkReviewPrev.checked = !!s.reviewPrevAfterEach;
  ui.checkTts.checked = !!s.ttsEnabled;
  populateTtsVoiceSelect();
  ui.checkAutoPlayNext.checked = !!s.autoPlayNextQa;
  if (ui.checkForceRecite) ui.checkForceRecite.checked = !!s.forceReciteCheck;
  if (ui.checkQaReciteSideBySide) ui.checkQaReciteSideBySide.checked = !!s.qaReciteSideBySide;
  ui.inputFullReadBefore.value = String(s.fullReadBeforeGroups || 1);
  ui.inputFullReadAfter.value = String(s.fullReadAfterGroups || 1);
  ui.inputReviewPrevRepeat.value = String(s.reviewPrevRepeatCount || 1);
  ui.inputRate.value = String(s.rate);
  ui.inputVolume.value = String(s.volume);
  ui.inputThreshold.value = String(s.threshold);
  ui.checkUseOllama.checked = !!s.useOllama;
  ui.inputOllamaUrl.value = s.ollamaUrl || '';
  ui.inputOllamaModel.value = s.ollamaModel || '';
  if (ui.checkReciteOnlyHighlights) ui.checkReciteOnlyHighlights.checked = !!s.reciteOnlyHighlights;
  updateManualCheckButton();
}

function renderCurrentQaView() {
  const qa = getCurrentQa();
  if (!qa) {
    ui.currentTitle.textContent = '未选择';
    ui.viewQuestion.textContent = '-';
    ui.viewAnswer.innerHTML = '';
    if (ui.btnMarkLevel0) ui.btnMarkLevel0.disabled = true;
    if (ui.btnMarkLevel1) ui.btnMarkLevel1.disabled = true;
    if (ui.btnMarkLevel2) ui.btnMarkLevel2.disabled = true;
    if (ui.reciteLevelStats) ui.reciteLevelStats.textContent = '0:0 1:0 2:0';
    renderQaTimer();
    return;
  }

  ui.currentTitle.textContent = qa.question || '（无问题）';
  ui.viewQuestion.textContent = qa.question || '-';
  const canMark = !!listState.collectionId;
  if (ui.btnMarkLevel0) ui.btnMarkLevel0.disabled = !canMark;
  if (ui.btnMarkLevel1) ui.btnMarkLevel1.disabled = !canMark;
  if (ui.btnMarkLevel2) ui.btnMarkLevel2.disabled = !canMark;
  if (ui.reciteLevelStats) {
    const counts = qa.reciteLevelCounts && typeof qa.reciteLevelCounts === 'object' ? qa.reciteLevelCounts : { 0: 0, 1: 0, 2: 0 };
    const c0 = Number(counts[0] ?? counts['0'] ?? 0) || 0;
    const c1 = Number(counts[1] ?? counts['1'] ?? 0) || 0;
    const c2 = Number(counts[2] ?? counts['2'] ?? 0) || 0;
    ui.reciteLevelStats.textContent = `0:${c0} 1:${c1} 2:${c2}`;
  }
  startQaTimerForCurrent();

  const sentences = splitSentences(qa.answerText);
  const activeIdx = player.qaId === qa.id ? player.activeSentenceGlobalIndex : null;
  const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
  const recited = ui.inputRecited.value || '';
  const onlyHighlights = !!state?.settings?.reciteOnlyHighlights;

  // When not in recite-check mode, prefer rich formatting preview
  // const hasCheckState = !!recited.trim() || reciteCheck.maskMode || reciteCheck.lockedSegments.size;
  const hasCheckState = reciteCheck.maskMode || reciteCheck.lockedSegments.size || reciteCheck.manualAuto;
  if (!cardCheck.enabled && !player.running && !hasCheckState && qa.answerHtml) {
    const safe = sanitizeAnswerHtml(normalizeImportedHtml(qa.answerHtml));
    ui.viewAnswer.innerHTML = `<div class="answer rich-answer rich-answer-editable">${safe || '<span class="muted">（无答案）</span>'}</div>`;
    return;
  }

  const { segments, bySentence } = getReciteSegmentsForQa(qa);
  const locked = reciteCheck.lockedSegments;

  if (cardCheck.enabled) {
    const idx = clamp(cardCheck.index, 0, Math.max(0, sentences.length - 1));
    cardCheck.index = idx;

    const segs = bySentence[idx] || [];
    const hit = segs.length ? segs.every((x) => locked.has(x.globalIndex)) : false;
    const masked = !cardCheck.flipped;
    const faceCls = ['card-face', masked ? 'masked' : '', hit ? 'hit' : recited.trim() ? 'miss' : '']
      .filter(Boolean)
      .join(' ');

    const content = sentences.length
      ? masked
        ? '（已遮住，点击“翻面”查看原文）'
        : escapeHtml(sentences[idx])
      : '（无答案）';

    const badge = sentences.length ? `第 ${idx + 1}/${sentences.length} 句` : '';
    const cardHtml = `
      <div class="card-area">
        <div class="${faceCls}">
          <div class="card-badge">${escapeHtml(badge)}</div>
          <div>${content}</div>
        </div>
      </div>
    `;

    ui.viewAnswer.innerHTML = cardHtml;
    ui.btnCardPrev.disabled = idx <= 0;
    ui.btnCardNext.disabled = idx >= sentences.length - 1;
    ui.btnCardFlip.disabled = !sentences.length;
    return;
  }

  if (onlyHighlights && qa.answerHtml) {
    const rawHtml = sanitizeAnswerHtml(normalizeImportedHtml(qa.answerHtml || ''));
    const doc = new DOMParser().parseFromString(rawHtml || '', 'text/html');
    let idx = 0;
    Array.from(doc.body?.querySelectorAll('mark') || []).forEach((m) => {
      const t = String(m.textContent || '').replace(/[\u00a0]/g, ' ').trim();
      if (!t) return;
      const isHit = locked.has(idx);
      const masked = reciteCheck.maskMode && !maskState.showAll && !isHit;
      const cls = ['sentence', 'segment', isHit ? 'hit' : '', masked ? 'masked' : ''].filter(Boolean).join(' ');

      const wrap = doc.createElement('span');
      wrap.setAttribute('class', cls);
      wrap.setAttribute('data-hseg', String(idx));

      const mk = doc.createElement('mark');
      const c = m.getAttribute('data-color');
      if (c) mk.setAttribute('data-color', c);
      if (masked) mk.textContent = maskPlaceholder(t);
      else mk.innerHTML = m.innerHTML;

      wrap.appendChild(mk);
      m.replaceWith(wrap);
      idx++;
    });

    ui.viewAnswer.innerHTML = `<div class="answer rich-answer">${doc.body?.innerHTML || ''}</div>`;

    ui.viewAnswer.querySelectorAll('[data-hseg]').forEach((node) => {
      node.addEventListener('click', async () => {
        const segIndex = Number(node.getAttribute('data-hseg'));
        if (!Number.isFinite(segIndex)) return;
        const seg = segments[segIndex];
        if (!seg) return;
        reciteCheck.lockedSegments.add(seg.globalIndex);
        renderCurrentQaView();
        await speak(seg.text);
        updateMatches();
      });
    });
    return;
  }

  const html = sentences
    .map((s, idx) => {
      const isActive = activeIdx === idx;
      const segs = bySentence[idx] || [];
      if (segs.length) {
        const segHtml = segs
          .map((seg) => {
            const isHit = locked.has(seg.globalIndex);
            const masked = reciteCheck.maskMode && !maskState.showAll && !isHit;
            const cls = ['sentence', 'segment', isActive ? 'active' : '', isHit ? 'hit' : '', masked ? 'masked' : '']
              .filter(Boolean)
              .join(' ');
            const display = masked ? maskPlaceholder(seg.text) : seg.text;
            return `<span class="${cls}" data-seg="${seg.globalIndex}">${escapeHtml(display)}</span>`;
          })
          .join('');
        return segHtml;
      }

      const sim = recited ? diceSimilarity(recited, s) : 0;
      const hasRecited = !!recited;
      const hit = hasRecited && sim >= threshold;
      const cls = ['sentence', isActive ? 'active' : '', hasRecited ? (hit ? 'hit' : 'miss') : ''].filter(Boolean).join(' ');
      const title = hasRecited ? `相似度：${sim.toFixed(2)}` : '';
      return `<span class="${cls}" data-idx="${idx}" title="${escapeHtml(title)}">${escapeHtml(s)}</span>`;
    })
    .join('');

  ui.viewAnswer.innerHTML = html || '<span class="muted">（无答案）</span>';

  ui.viewAnswer.querySelectorAll('[data-seg]').forEach((node) => {
    node.addEventListener('click', async () => {
      const segIndex = Number(node.getAttribute('data-seg'));
      if (!Number.isFinite(segIndex)) return;
      const seg = segments[segIndex];
      if (!seg) return;

      reciteCheck.lockedSegments.add(seg.globalIndex);
      renderCurrentQaView();
      await speak(seg.text);
      updateMatches();
    });
  });
}

function fillEditorFromCurrent() {
  const qa = getCurrentQa();
  if (!qa) return;
  ui.inputQuestion.value = qa.question || '';
  if (ui.inputAnswerRich) {
    if (qa.answerHtml) {
      setRichEditorHtml(qa.answerHtml);
    } else {
      // Backward compatible: show plain text
      setRichEditorHtml(escapeHtml(qa.answerText || '').replace(/\n/g, '<br>'));
    }
  }
  ui.inputAnswer.value = qa.answerText || '';
  if (ui.inputFocusPoints) ui.inputFocusPoints.value = qa.focusPoints || '';
}

function renderSpeechAvailability() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    ui.speechHint.textContent = '当前浏览器不支持语音识别。请使用 Chrome/Edge。你仍然可以手动输入背诵内容进行检测。';
    ui.btnRecStart.disabled = true;
    ui.btnRecStop.disabled = true;
    return;
  }
  ui.speechHint.textContent = '提示：建议使用 Chrome/Edge。首次录音可能会弹出麦克风权限请求。';
  ui.btnRecStart.disabled = false;
  ui.btnRecStop.disabled = false;
}

function updateCardButtons() {
  ui.btnCardMode.textContent = cardCheck.enabled ? '卡片检查：开' : '卡片检查：关';
  ui.btnMaskToggleAll.textContent = maskState.showAll ? '重新遮住' : '全部显示';
}

function renderStatus() {
  if (!player.running) {
    ui.statusLine.textContent = '请选择一个 QA，然后点击开始。';
    return;
  }
  const step = player.steps[player.stepIndex];
  if (!step) {
    ui.statusLine.textContent = '已完成。';
    return;
  }
  if (step.isFullRead) {
    const prefix = step.fullReadType === 'before' ? '分组前完整朗读：' : '分组后完整朗读：';
    const type = step.isQuestion ? '问题' : '答案';
    ui.statusLine.textContent = `${prefix}${type}：${step.text}`;
  } else if (step.isQuestion) {
    const prefix = step.review ? '回顾上一对：问题：' : '问题：';
    ui.statusLine.textContent = `${prefix}${step.text}`;
  } else {
    const prefix = player.reviewMode ? '回顾上一对：' : '';
    ui.statusLine.textContent = `${prefix}组 ${step.groupIndex + 1}/${step.groupCount}，第 ${step.round + 1}/${step.roundCount} 遍：${step.text}`;
  }
}

function render() {
  renderQaList();
  renderSettings();
  updateCardButtons();
  renderCurrentQaView();
  renderSpeechAvailability();
  renderQaTimer();

  if (ui.qaReciteWrap) {
    ui.qaReciteWrap.classList.toggle('is-side-by-side', !!state?.settings?.qaReciteSideBySide);
  }
}

function deleteSelectedQas() {
  if (!listState.selected.size) return;
  const ok = confirm(`确认删除选中的 ${listState.selected.size} 条 QA？`);
  if (!ok) return;

  stopPlayer();
  const selected = new Set(listState.selected);
  state.qas = state.qas.filter((q) => !selected.has(q.id));
  listState.selected.clear();

  if (state.progress.currentQaId && selected.has(state.progress.currentQaId)) {
    state.progress.currentQaId = state.qas.length ? state.qas[0].id : null;
  }

  saveState();
  resetReciteCheck();
  resetCardCheck();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function upsertQa(qa) {
  const idx = state.qas.findIndex((x) => x.id === qa.id);
  if (idx >= 0) state.qas[idx] = qa;
  else state.qas.unshift(qa);
  saveState();
}

function deleteCurrentQa() {
  const id = state.progress.currentQaId;
  if (!id) return;
  const idx = state.qas.findIndex((x) => x.id === id);
  if (idx < 0) return;
  state.qas.splice(idx, 1);
  if (state.qas.length) state.progress.currentQaId = state.qas[Math.min(idx, state.qas.length - 1)].id;
  else state.progress.currentQaId = null;
  saveState();
}

function applySettingsFromUI() {
  state.settings.groupSize = clamp(parseInt(ui.inputGroupSize.value || '3', 10), 1, 10);
  state.settings.repeatPerGroup = clamp(parseInt(ui.inputRepeat.value || '4', 10), 1, 20);
  state.settings.sentenceDelimiters = String(ui.inputSentenceDelims.value || '').replace(/\s+/g, '');
  state.settings.reviewPrevAfterEach = !!ui.checkReviewPrev.checked;
  state.settings.ttsEnabled = !!ui.checkTts.checked;
  state.settings.ttsVoiceUri = String(ui.selectTtsVoice?.value || '').trim();
  state.settings.autoPlayNextQa = !!ui.checkAutoPlayNext.checked;
  state.settings.forceReciteCheck = !!ui.checkForceRecite?.checked;
  state.settings.qaReciteSideBySide = !!ui.checkQaReciteSideBySide?.checked;
  state.settings.reciteOnlyHighlights = !!ui.checkReciteOnlyHighlights?.checked;
  state.settings.fullReadBeforeGroups = clamp(parseInt(ui.inputFullReadBefore.value || '1', 10), 0, 5);
  state.settings.fullReadAfterGroups = clamp(parseInt(ui.inputFullReadAfter.value || '1', 10), 0, 5);
  state.settings.reviewPrevRepeatCount = clamp(parseInt(ui.inputReviewPrevRepeat.value || '1', 10), 1, 5);
  state.settings.rate = clamp(Number(ui.inputRate.value || '1'), 0.5, 4);
  state.settings.volume = clamp(Number(ui.inputVolume.value || '1'), 0, 1);
  state.settings.threshold = clamp(Number(ui.inputThreshold.value || '0.65'), 0, 1);
  state.settings.useOllama = !!ui.checkUseOllama.checked;
  state.settings.ollamaUrl = ui.inputOllamaUrl.value || '';
  state.settings.ollamaModel = ui.inputOllamaModel.value || '';
  saveState();
}

function stopTts() {
  try {
    speechSynthesis.cancel();
  } catch {}
  player.speechUtterance = null;
}

function isReciteCheckPassedForQa(qa) {
  if (!qa) return false;
  const { segments } = getReciteSegmentsForQa(qa);
  if (!segments.length) return true;
  const hits = [...reciteCheck.lockedSegments].filter((i) => i >= 0 && i < segments.length).length;
  return hits >= segments.length;
}

function warnIfReciteNotPassed() {
  if (!state.settings.forceReciteCheck) return;
  const qa = getCurrentQa();
  if (!qa) return;
  if (isReciteCheckPassedForQa(qa)) return;
  ui.statusLine.textContent = '强制背诵检查：当前题未完全命中（自动下一题已禁用）。你仍可手动进入下一题。';
}

function pickTtsVoice() {
  const voices = listVoices();
  const desired = String(state?.settings?.ttsVoiceUri || '').trim();
  if (desired) {
    const v = voices.find((x) => (x.voiceURI || '') === desired);
    if (v) return v;
  }
  const zh = voices.find((v) => (v.lang || '').toLowerCase().startsWith('zh'));
  return zh || voices[0] || null;
}

function voicesCount() {
  try {
    return (speechSynthesis.getVoices?.() || []).length;
  } catch {
    return 0;
  }
}

function speak(text) {
  const t = (text || '').trim();
  if (!t) return Promise.resolve();

  if (!state.settings.ttsEnabled) {
    ui.statusLine.textContent = '未朗读：你关闭了“朗读文字（TTS）”。';
    return Promise.resolve();
  }
  if (!('speechSynthesis' in window)) {
    ui.statusLine.textContent = '未朗读：当前浏览器不支持语音合成（TTS）。请使用 Chrome/Edge。';
    return Promise.resolve();
  }

  const splitForTts = (raw) => {
    const s = String(raw || '').trim();
    if (!s) return [];
    const maxLen = 160;
    if (s.length <= maxLen) return [s];
    const parts = [];
    let buf = '';
    for (const ch of s) {
      buf += ch;
      const hit = /[。！？!?；;，,\n]/.test(ch);
      if (buf.length >= maxLen || (hit && buf.length >= 60)) {
        const p = buf.trim();
        if (p) parts.push(p);
        buf = '';
      }
    }
    if (buf.trim()) parts.push(buf.trim());
    return parts.length ? parts : [s];
  };

  const isAndroid = /Android/i.test(navigator.userAgent || '');

  const speakOnce = (piece) =>
    new Promise((resolve) => {
      const p = (piece || '').trim();
      if (!p) return resolve();

      try {
        speechSynthesis.getVoices?.();
      } catch {}

      // Android often gets speechSynthesis into a stuck speaking/pending state.
      // These heuristics can cause false negatives on desktop, so only enable on Android.
      const waitIdleMaxMs = isAndroid ? 800 : 0;
      const idleStart = Date.now();

      const isBusy = () => {
        try {
          return !!(speechSynthesis.speaking || speechSynthesis.pending);
        } catch {
          return false;
        }
      };

      const waitForIdleOrCancel = () => {
        if (!isBusy()) return true;
        if (waitIdleMaxMs > 0 && Date.now() - idleStart >= waitIdleMaxMs) {
          try {
            stopTts();
          } catch {}
          return true;
        }
        return waitIdleMaxMs === 0;
      };

      const attemptSpeak = (attempt) => {
        if (!waitForIdleOrCancel()) {
          setTimeout(() => attemptSpeak(attempt), 30);
          return;
        }

        const u = new SpeechSynthesisUtterance(p);
        u.lang = 'zh-CN';
        const voice = pickTtsVoice();
        if (voice) u.voice = voice;
        u.rate = clamp(Number(state.settings.rate) || 1, 0.5, 4);
        u.volume = clamp(Number(state.settings.volume) || 1, 0, 1);

        let done = false;
        let started = false;
        const finish = () => {
          if (done) return;
          done = true;
          resolve();
        };

        const estMs = Math.round((p.length * 450) / Math.max(0.1, u.rate) + 2000);
        const watchdogMs = clamp(estMs, 5000, 300000);
        const watchdog = setTimeout(() => {
          ui.statusLine.textContent = `提示：本句 TTS 超时（已自动继续下一句）。voices=${voicesCount()}`;
          finish();
        }, watchdogMs);

        const startWatchdogMs = isAndroid ? 3500 : 0;
        const startWatchdog = startWatchdogMs
          ? setTimeout(() => {
              // Some browsers/voices don't fire onstart reliably; also treat "becoming busy" as started.
              if (started) return;
              if (isBusy()) {
                started = true;
                return;
              }
              // If it never started, cancel and retry once (Android only).
              try {
                stopTts();
              } catch {}
              clearTimeout(watchdog);
              if (attempt < 1) {
                setTimeout(() => attemptSpeak(attempt + 1), 120);
              } else {
                ui.statusLine.textContent = `未朗读：TTS 启动失败（已跳过）。voices=${voicesCount()}`;
                finish();
              }
            }, startWatchdogMs)
          : null;

        u.onstart = () => {
          started = true;
        };
        u.onend = () => {
          if (startWatchdog) clearTimeout(startWatchdog);
          clearTimeout(watchdog);
          finish();
        };
        u.onerror = () => {
          if (startWatchdog) clearTimeout(startWatchdog);
          clearTimeout(watchdog);
          finish();
        };
        player.speechUtterance = u;

        try {
          speechSynthesis.speak(u);
          // Some Android builds require an explicit resume() after speak().
          speechSynthesis.resume?.();
        } catch {
          if (startWatchdog) clearTimeout(startWatchdog);
          clearTimeout(watchdog);
          if (attempt < 1) {
            setTimeout(() => attemptSpeak(attempt + 1), 80);
          } else {
            ui.statusLine.textContent = '未朗读：speechSynthesis.speak 调用失败。';
            finish();
          }
        }
      };

      attemptSpeak(0);
    });

  const pieces = splitForTts(t);
  return pieces.reduce((acc, piece) => acc.then(() => speakOnce(piece)), Promise.resolve());
}

function buildStepsForQa(qa, opts) {
  const sentences = splitSentences(qa.answerText);
  const groupSize = clamp(Number(state.settings.groupSize) || 3, 1, 10);
  const roundCount = clamp(Number(state.settings.repeatPerGroup) || 4, 1, 20);
  const groups = chunk(sentences, groupSize);

  const steps = [];
  for (let g = 0; g < groups.length; g++) {
    for (let r = 0; r < roundCount; r++) {
      for (let s = 0; s < groups[g].length; s++) {
        const globalIdx = g * groupSize + s;
        steps.push({
          qaId: qa.id,
          groupIndex: g,
          groupCount: groups.length,
          round: r,
          roundCount,
          sentenceIndexInGroup: s,
          globalSentenceIndex: globalIdx,
          text: groups[g][s],
          kind: 'speak',
          review: !!opts.review,
        });
      }
    }
  }

  return steps;
}

function prevQaId(currentId) {
  const idx = state.qas.findIndex((q) => q.id === currentId);
  if (idx <= 0) return null;
  return state.qas[idx - 1].id;
}

function nextQaId(currentId) {
  const idx = state.qas.findIndex((q) => q.id === currentId);
  if (idx < 0 || idx >= state.qas.length - 1) return null;
  return state.qas[idx + 1].id;
}

async function runPlayerLoop() {
  while (player.running) {
    if (!state.settings.ttsEnabled) {
      player.running = false;
      player.paused = false;
      player.activeSentenceGlobalIndex = null;
      stopTts();
      render();
      ui.statusLine.textContent = '已停止：你关闭了“朗读文字（TTS）”。';
      return;
    }

    if (player.paused) {
      await new Promise((r) => setTimeout(r, 80));
      continue;
    }

    const step = player.steps[player.stepIndex];
    if (!step) {
      const finishedQaId = player.mainQaId || player.qaId;
      // If force recite check is enabled, do NOT auto-advance. User must pass check then click next.
      const nextId = (!state.settings.forceReciteCheck && state.settings.autoPlayNextQa && finishedQaId)
        ? nextQaId(finishedQaId)
        : null;
      if (nextId) {
        const nextQa = getQaById(nextId);
        if (!nextQa) {
          player.running = false;
          player.paused = false;
          player.activeSentenceGlobalIndex = null;
          render();
          return;
        }

        if ((listState.query || '').trim()) {
          listState.query = '';
          ui.inputQaSearch.value = '';
        }
        const idxInAll = state.qas.findIndex((x) => x.id === nextId);
        if (idxInAll >= 0) listState.page = Math.floor(idxInAll / listState.pageSize) + 1;

        state.progress.currentQaId = nextId;
        saveState();
        ui.inputRecited.value = '';
        resetReciteCheck();
        resetCardCheck();

        player.paused = false;
        player.qaId = nextId;
        player.mainQaId = nextId;
        player.steps = buildPlayerStepsForQa(nextQa);
        player.stepIndex = 0;
        player.activeSentenceGlobalIndex = null;
        player.reviewMode = false;

        render();
        updateMatches();
        continue;
      }

      if (state.settings.forceReciteCheck) {
        const qa = getCurrentQa();
        if (qa && isReciteCheckPassedForQa(qa)) {
          ui.statusLine.textContent = '本题播放完成：背诵检查已通过。请手动点击“下一题”进入下一题。';
        } else {
          ui.statusLine.textContent = '本题播放完成：请先完成背诵检查（命中全部段落），否则无法进入下一题。';
        }
      }

      player.running = false;
      player.paused = false;
      player.activeSentenceGlobalIndex = null;
      render();
      return;
    }

    const qa = getQaById(step.qaId);
    if (!qa) {
      player.stepIndex++;
      continue;
    }

    player.qaId = qa.id;
    player.activeSentenceGlobalIndex = step.globalSentenceIndex;
    player.reviewMode = !!step.review;

    render();

    const startIndex = player.stepIndex;
    await speak(step.text);
    // If paused/stopped/canceled during this speak, do not advance automatically.
    if (!player.running) return;
    if (player.paused) continue;
    if (player.stepIndex !== startIndex) continue;
    player.stepIndex++;
  }
}

function startPlayer() {
  if (!state.settings.ttsEnabled) {
    ui.statusLine.textContent = '无法开始：请先开启“朗读文字（TTS）”。';
    return;
  }
  let qa = getCurrentQa();
  if (!qa && state.qas.length) {
    state.progress.currentQaId = state.qas[0].id;
    saveState();
    qa = getCurrentQa();
  }
  if (!qa) {
    ui.statusLine.textContent = '无法开始：请先选择或创建一个 QA。';
    return;
  }
  player.running = true;
  player.paused = false;
  player.qaId = qa.id;
  player.mainQaId = qa.id;
  player.stepIndex = 0;
  player.activeSentenceGlobalIndex = null;
  player.reviewMode = false;

  player.steps = buildPlayerStepsForQa(qa);

  runPlayerLoop();
}

function pausePlayer() {
  if (!player.running) return;
  player.paused = true;
  stopTts();
  render();
}

function resumePlayer() {
  if (!player.running) return;
  if (!state.settings.ttsEnabled) {
    ui.statusLine.textContent = '无法继续：请先开启“朗读文字（TTS）”。';
    return;
  }
  player.paused = false;
  render();
}

function stopPlayer() {
  player.running = false;
  player.paused = false;
  player.steps = [];
  player.stepIndex = 0;
  player.activeSentenceGlobalIndex = null;
  player.qaId = null;
  player.mainQaId = null;
  player.reviewMode = false;
  stopTts();
  render();
}

function startPlayerFromSentence(qa, sentenceIndex) {
  if (!qa) return;
  const steps = buildPlayerStepsForQa(qa);
  const idx = steps.findIndex((s) => s.kind === 'speak' && s.globalSentenceIndex === sentenceIndex && !s.review);
  player.running = true;
  player.paused = false;
  player.qaId = qa.id;
  player.steps = steps;
  player.stepIndex = idx >= 0 ? idx : 0;
  player.activeSentenceGlobalIndex = null;
  player.reviewMode = false;
  runPlayerLoop();
}

function stepSentence(delta) {
  // Only meaningful when a QA is selected
  const qa = getCurrentQa();
  if (!qa) return;
  // If currently playing, move to nearest next/prev speak step within current QA
  if (player.running) {
    const start = clamp(player.stepIndex + delta, 0, Math.max(0, player.steps.length - 1));
    let i = start;
    while (i >= 0 && i < player.steps.length) {
      const st = player.steps[i];
      if (st && st.kind === 'speak' && !st.isQuestion) {
        player.stepIndex = i;
        player.activeSentenceGlobalIndex = st.globalSentenceIndex ?? null;
        render();
        return;
      }
      i += delta >= 0 ? 1 : -1;
    }
    return;
  }

  // Not playing: control card mode sentence index as "上一句/下一句"
  if (!cardCheck.enabled) cardCheck.enabled = true;
  const total = splitSentences(qa.answerText).length;
  if (!total) return;
  cardCheck.index = clamp(cardCheck.index + delta, 0, total - 1);
  cardCheck.flipped = false;
  renderCurrentQaView();
}

function gotoPrevQa() {
  const cur = state.progress.currentQaId;
  if (!cur) return;
  const id = prevQaId(cur);
  if (!id) return;
  stopPlayer();
  setCurrentQa(id);
  fillEditorFromCurrent();
}

function gotoNextQa() {
  const cur = state.progress.currentQaId;
  if (!cur) return;
  warnIfReciteNotPassed();
  const id = nextQaId(cur);
  if (!id) return;
  stopPlayer();
  setCurrentQa(id);
  fillEditorFromCurrent();
}

let recognition = null;

function ensureRecognition() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) return null;
  if (recognition) return recognition;

  const r = new SpeechRecognition();
  r.lang = 'zh-CN';
  r.continuous = true;
  r.interimResults = true;

  let interim = '';

  r.onresult = (event) => {
    interim = '';
    for (let i = event.resultIndex; i < event.results.length; i++) {
      const res = event.results[i];
      const text = (res[0]?.transcript || '').trim();
      if (!text) continue;
      if (res.isFinal) {
        handleFinalUtterance(text);
      } else {
        interim += text;
      }
    }
    if (interim) ui.speechHint.textContent = `识别中：${interim}`;
  };

  r.onerror = () => {
    ui.speechHint.textContent = '语音识别出现问题。你可以改用手动输入检测。';
  };

  recognition = r;
  return r;
}

function startRecording() {
  const r = ensureRecognition();
  if (!r) return;
  try {
    ui.speechHint.textContent = '正在录音识别中…';
    resetReciteCheck();
    ui.inputRecited.value = '';
    reciteCheck.maskMode = true;
    const qa = getCurrentQa();
    if (qa) {
      const { segments } = getReciteSegmentsForQa(qa);
      advanceSegmentPointer(segments);
    }
    renderCurrentQaView();
    r.start();
  } catch {
    ui.speechHint.textContent = '录音启动失败。请检查麦克风权限或稍后重试。';
  }
}

function stopRecording() {
  // if (!recognition) return;
  // try {
  //   recognition.stop();
  //   ui.speechHint.textContent = '已停止录音。你可以继续编辑识别文本并检测。';
  // } catch {}
  const hadRecognition = !!recognition;
  try {
    if (recognition) recognition.stop();
    ui.speechHint.textContent = hadRecognition ? '已停止录音。' : '已停止录音。';
  } catch {}
  resetReciteCheck();
  render();
  updateMatches();
}

function advanceSegmentPointer(segments) {
  if (reciteCheck.pointerSegment < 0) reciteCheck.pointerSegment = 0;
  while (reciteCheck.pointerSegment < segments.length && reciteCheck.lockedSegments.has(reciteCheck.pointerSegment)) {
    reciteCheck.pointerSegment++;
  }
  if (cardCheck.enabled) {
    const current = segments[reciteCheck.pointerSegment];
    if (current) {
      cardCheck.index = clamp(current.sentenceIndex, 0, Math.max(0, splitSentences(getCurrentQa()?.answerText).length - 1));
      cardCheck.flipped = false;
    }
  }
}

function handleFinalUtterance(text) {
  const qa = getCurrentQa();
  if (!qa) return;

  const { segments, bySentence } = getReciteSegmentsForQa(qa);
  reciteCheck.pointerSegment = 0;
  advanceSegmentPointer(segments);

  if (reciteCheck.pointerSegment >= segments.length) {
    ui.speechHint.textContent = '已全部命中。';
    return;
  }

  const targetSeg = segments[reciteCheck.pointerSegment];
  const target = targetSeg?.text || '';
  const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
  const sim = diceSimilarity(text, target);

  reciteCheck.lastUtterance = text;
  ui.inputRecited.value = (ui.inputRecited.value ? ui.inputRecited.value.trimEnd() + '\n' : '') + text;

  let best = null;
  for (const seg of segments) {
    if (reciteCheck.lockedSegments.has(seg.globalIndex)) continue;
    const s = diceSimilarity(text, seg.text);
    if (!best || s > best.sim) best = { seg, sim: s };
  }

  if (best && best.sim >= threshold) {
    reciteCheck.lockedSegments.add(best.seg.globalIndex);
    reciteCheck.pointerSegment = 0;
    advanceSegmentPointer(segments);
    const sentenceSegs = bySentence[best.seg.sentenceIndex] || [];
    const sentenceDone = sentenceSegs.length ? sentenceSegs.every((x) => reciteCheck.lockedSegments.has(x.globalIndex)) : false;
    ui.speechHint.textContent = sentenceDone
      ? `命中：第 ${best.seg.sentenceIndex + 1} 句完成（${best.sim.toFixed(2)}）`
      : `命中：第 ${best.seg.sentenceIndex + 1} 句第 ${best.seg.segmentIndex + 1} 段（${best.sim.toFixed(2)}）`;
  } else {
    ui.speechHint.textContent = best
      ? `未命中（最高 ${best.sim.toFixed(2)} < 阈值）`
      : '未命中';
  }

  updateMatches();
}

function computeMatch(qa, recitedText) {
  const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
  const onlyHi = !!state?.settings?.reciteOnlyHighlights;
  if (onlyHi && qa?.answerHtml) {
    const parts = extractHighlightSegmentsFromAnswerHtml(qa.answerHtml);
    const scores = parts.map((t) => diceSimilarity(recitedText, t));
    const hits = scores.filter((x) => x >= threshold).length;
    const total = parts.length;
    const pct = total ? Math.round((hits / total) * 100) : 0;
    const min = scores.length ? Math.min(...scores) : 0;
    const max = scores.length ? Math.max(...scores) : 0;
    return { threshold, hits, total, pct, min, max, mode: 'highlight' };
  }

  const sentences = splitSentences(qa?.answerText || '');
  const scores = sentences.map((s) => diceSimilarity(recitedText, s));
  const hits = scores.filter((x) => x >= threshold).length;
  const total = sentences.length;
  const pct = total ? Math.round((hits / total) * 100) : 0;
  const min = scores.length ? Math.min(...scores) : 0;
  const max = scores.length ? Math.max(...scores) : 0;
  return { threshold, hits, total, pct, min, max, mode: 'sentence' };
}

function updateMatches() {
  const qa = getCurrentQa();
  if (!qa) {
    ui.matchSummary.textContent = '-';
    return;
  }
  const recited = ui.inputRecited.value || '';
  if (!recited.trim()) {
    ui.matchSummary.textContent = '输入或录音后会显示命中情况。';
    renderCurrentQaView();
    return;
  }

  if (reciteCheck.lockedSegments.size || reciteCheck.maskMode) {
    const { segments, bySentence } = getReciteSegmentsForQa(qa);

    const segHits = [...reciteCheck.lockedSegments].filter((i) => i >= 0 && i < segments.length).length;
    const segTotal = segments.length;

    const sentenceHits = bySentence.filter((segs) => segs && segs.length && segs.every((x) => reciteCheck.lockedSegments.has(x.globalIndex))).length;
    const sentenceTotal = bySentence.length;

    const pct = segTotal ? Math.round((segHits / segTotal) * 100) : 0;
    const nextSeg = Math.min(reciteCheck.pointerSegment + 1, segTotal);
    ui.matchSummary.textContent = `命中 ${segHits}/${segTotal} 段（${pct}%），完成 ${sentenceHits}/${sentenceTotal} 句，阈值 ${clamp(Number(state.settings.threshold) || 0.65, 0, 1).toFixed(2)}，下一段 ${nextSeg}/${segTotal}`;
    renderCurrentQaView();
    return;
  }

  const m = computeMatch(qa, recited);
  const unit = m.mode === 'highlight' ? '段（高亮）' : '句';
  ui.matchSummary.textContent = `命中 ${m.hits}/${m.total} ${unit}（${m.pct}%），阈值 ${m.threshold.toFixed(2)}，相似度范围 ${m.min.toFixed(2)}~${m.max.toFixed(2)}`;
  renderCurrentQaView();
}

function runManualCheckFromText(raw, preserveAuto) {
  const qa = getCurrentQa();
  if (!qa) return;
  const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
  const { segments } = getReciteSegmentsForQa(qa);

  const prevAuto = reciteCheck.manualAuto;
  resetReciteCheck();
  if (preserveAuto) reciteCheck.manualAuto = prevAuto;
  updateManualCheckButton();

  if (!raw.trim()) {
    updateMatches();
    return;
  }

  reciteCheck.maskMode = true;
  reciteCheck.pointerSegment = 0;
  advanceSegmentPointer(segments);

  let utterances = raw
    .split(/\n+/g)
    .map((x) => x.trim())
    .filter(Boolean);

  if (utterances.length === 1 && utterances[0].length > 60) {
    const extra = '，,;；';
    const delims = String(state.settings.sentenceDelimiters || '。！？!?') + extra;
    const cls = delims.replace(/[\\\]\[\-\^]/g, (m) => `\\${m}`);
    try {
      utterances = utterances[0]
        .split(new RegExp(`[${cls}]`, 'g'))
        .map((x) => x.trim())
        .filter(Boolean);
    } catch {
      // keep as is
    }
  }

  let hitCount = 0;
  for (const u of utterances) {
    let best = null;
    for (const seg of segments) {
      if (reciteCheck.lockedSegments.has(seg.globalIndex)) continue;
      const s = diceSimilarity(u, seg.text);
      if (!best || s > best.sim) best = { seg, sim: s };
    }

    if (best && best.sim >= threshold) {
      reciteCheck.lockedSegments.add(best.seg.globalIndex);
      hitCount++;
      reciteCheck.pointerSegment = 0;
      advanceSegmentPointer(segments);
    }
  }

  ui.speechHint.textContent = hitCount
    ? `输入检测完成：命中 ${hitCount} 段。`
    : `输入检测完成：未命中（阈值 ${threshold.toFixed(2)}）。`;
  updateMatches();
}

function applyManualCheck(preserveAuto) {
  runManualCheckFromText(ui.inputRecited.value || '', !!preserveAuto);
}

function stripHtmlToText(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return (doc.body?.textContent || '').replace(/[\u00a0]/g, ' ').replace(/[ \t]+/g, ' ').trim();
}

function parseDocxHtmlToQas(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const body = doc.body;
  if (!body) return [];

  const qas = [];
  let curTitle = '';
  let curParts = [];
  let curFocusLines = [];
  let focusMode = false;

  const flush = () => {
    const question = (curTitle || '').trim();
    const answerHtml = curParts.join('').trim();
    const answerText = htmlToPlainTextPreserveLines(answerHtml);
    const focusPoints = curFocusLines.map((x) => String(x || '').trim()).filter(Boolean).join('\n');
    if (question || answerText) {
      qas.push({
        id: uid(),
        question: question || '（未命名）',
        answerHtml: sanitizeAnswerHtml(normalizeImportedHtml(answerHtml)),
        answerText,
        focusPoints,
        createdAt: nowIso(),
        updatedAt: nowIso(),
      });
    }
    curTitle = '';
    curParts = [];
    curFocusLines = [];
    focusMode = false;
  };

  const nodes = Array.from(body.children);
  for (const n of nodes) {
    const tag = (n.tagName || '').toUpperCase();
    const text = (n.textContent || '').replace(/[\u00a0]/g, ' ').trim();
    if (!text) continue;

    const isHeading = /^H[1-6]$/.test(tag);
    if (isHeading) {
      const level = Number(tag.slice(1));
      // Our exported format uses H3 as focus-points marker.
      if (level === 3 && text === '背诵需关注的点') {
        focusMode = true;
        continue;
      }

      // If we already started a QA (has title/answer/focus), a new heading (except focus marker) means flush.
      if (curTitle || curParts.length || curFocusLines.length) flush();

      // Any heading other than the focus marker is treated as a QA question.
      curTitle = text;
      focusMode = false;
      continue;
    }

    if (!curTitle && !qas.length) {
      curTitle = text;
      continue;
    }

    if (focusMode) {
      curFocusLines.push(text);
      continue;
    }

    // Keep formatting from mammoth HTML
    curParts.push(`<p>${normalizeImportedHtml(n.innerHTML || '')}</p>`);
  }
  if (curTitle || curParts.length) flush();
  return qas;
}

async function docxArrayBufferToHtml(buf) {
  const JSZipLib = window.JSZip;
  if (!JSZipLib) return '';

  const zip = await JSZipLib.loadAsync(buf);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return '';
  const documentXml = await docFile.async('string');
  if (!documentXml) return '';

  const stylesFile = zip.file('word/styles.xml');
  const stylesXml = stylesFile ? await stylesFile.async('string') : '';

  const parseStyles = (xml) => {
    const map = new Map();
    if (!xml) return map;
    const d = new DOMParser().parseFromString(String(xml), 'application/xml');
    const styles = Array.from(d.getElementsByTagName('w:style') || []);
    styles.forEach((s) => {
      const type = s.getAttribute('w:type') || s.getAttribute('type');
      if (type && String(type).toLowerCase() !== 'paragraph') return;
      const id = s.getAttribute('w:styleId') || s.getAttribute('styleId');
      if (!id) return;
      const n = s.getElementsByTagName('w:name')?.[0];
      const name = n?.getAttribute('w:val') || n?.getAttribute('val') || '';
      if (name) map.set(String(id), String(name));
    });
    return map;
  };

  const styleMap = parseStyles(stylesXml);

  const getLocal = (node) => String(node?.tagName || '').split(':').pop().toLowerCase();

  const highlightToColor = (val) => {
    const v = String(val || '').toLowerCase();
    if (!v || v === 'none') return '';
    if (v.includes('yellow')) return 'yellow';
    if (v.includes('green')) return 'green';
    if (v.includes('cyan') || v.includes('blue')) return 'cyan';
    if (v.includes('magenta') || v.includes('red') || v.includes('pink')) return 'magenta';
    if (v.includes('darkcyan') || v.includes('darkblue')) return 'cyan';
    if (v.includes('darkgreen')) return 'green';
    if (v.includes('darkmagenta') || v.includes('darkred')) return 'magenta';
    if (v.includes('darkyellow')) return 'yellow';
    return 'yellow';
  };

  const getHeadingLevel = (styleId) => {
    const sid = String(styleId || '');
    const name = styleMap.get(sid) || '';
    const s = `${sid} ${name}`.toLowerCase();
    const m1 = s.match(/heading\s*([1-6])/i);
    if (m1) return Number(m1[1]);
    const m2 = s.match(/标题\s*([1-6])/i);
    if (m2) return Number(m2[1]);
    const m3 = s.match(/heading([1-6])/i);
    if (m3) return Number(m3[1]);
    return 0;
  };

  const buildPlainText = (p) => {
    const ts = Array.from(p.getElementsByTagName('w:t') || []);
    return ts.map((x) => x.textContent || '').join('');
  };

  const buildRichHtml = (p) => {
    const out = [];

    const handleRun = (r) => {
      const rPr = r.getElementsByTagName('w:rPr')?.[0];
      const bold = !!(rPr && (rPr.getElementsByTagName('w:b')?.[0] || rPr.getElementsByTagName('w:bCs')?.[0]));
      const hl = rPr?.getElementsByTagName('w:highlight')?.[0];
      const hlVal = hl ? (hl.getAttribute('w:val') || hl.getAttribute('val')) : '';
      const color = highlightToColor(hlVal);

      const parts = [];
      Array.from(r.childNodes || []).forEach((c) => {
        const l = getLocal(c);
        if (l === 't') {
          const t = c.textContent || '';
          if (t) parts.push(escapeHtml(t));
        } else if (l === 'tab') {
          parts.push('    ');
        } else if (l === 'br' || l === 'cr') {
          parts.push('<br>');
        }
      });

      const inner = parts.join('');
      if (!inner) return;

      let wrapped = inner;
      if (bold) wrapped = `<strong>${wrapped}</strong>`;
      if (color) wrapped = `<mark data-color="${color}">${wrapped}</mark>`;
      out.push(wrapped);
    };

    const walk = (node) => {
      const l = getLocal(node);
      if (l === 'r') return handleRun(node);
      Array.from(node.childNodes || []).forEach((c) => walk(c));
    };

    walk(p);
    return out.join('');
  };

  const xmlDoc = new DOMParser().parseFromString(String(documentXml), 'application/xml');
  const ps = Array.from(xmlDoc.getElementsByTagName('w:p') || []);
  if (!ps.length) return '';

  const blocks = [];
  ps.forEach((p) => {
    const pStyle = p.getElementsByTagName('w:pStyle')?.[0];
    const styleId = pStyle ? (pStyle.getAttribute('w:val') || pStyle.getAttribute('val')) : '';
    const level = getHeadingLevel(styleId);
    if (level) {
      const text = buildPlainText(p).replace(/[\u00a0]/g, ' ').trim();
      if (text) blocks.push(`<h${level}>${escapeHtml(text)}</h${level}>`);
      return;
    }

    const rich = buildRichHtml(p).trim();
    const plain = buildPlainText(p).replace(/[\u00a0]/g, ' ').trim();
    if (!rich && !plain) return;
    blocks.push(`<p>${rich || escapeHtml(plain)}</p>`);
  });

  return blocks.join('');
}

async function importDocx(file) {
  if (!window.mammoth) {
    alert('docx 导入需要联网加载 mammoth 库（CDN）。');
    return;
  }
  const buf = await file.arrayBuffer();
  let html = '';
  try {
    html = (await docxArrayBufferToHtml(buf)) || '';
  } catch {
    html = '';
  }
  if (!html) {
    const result = await window.mammoth.convertToHtml({ arrayBuffer: buf });
    html = result?.value || '';
  }
  const qas = parseDocxHtmlToQas(html);
  if (!qas.length) {
    const text = stripHtmlToText(html);
    if (text) {
      qas.push({
        id: uid(),
        question: file.name.replace(/\.docx$/i, ''),
        answerHtml: sanitizeAnswerHtml(normalizeImportedHtml(html)),
        answerText: text,
        createdAt: nowIso(),
        updatedAt: nowIso(),
      });
    }
  }
  if (!qas.length) {
    alert('未从 docx 中解析出内容。建议用“标题(Heading)+正文段落”的格式。');
    return;
  }

  const col = createCollection(normalizeImportedCollectionName(file?.name));
  const attached = qas.map((q) => ({ ...q, id: uid(), collectionId: col.id }));
  state.qas = [...attached, ...state.qas];
  state.progress.currentQaId = attached[0].id;
  setActiveCollection(col.id);
  listState.page = 1;
  listState.selected.clear();
  saveState();
  stopPlayer();
  resetReciteCheck();
  resetCardCheck();
  render();
  fillEditorFromCurrent();
  updateMatches();
}

function offlineAnswer(qa, question) {
  const sentences = splitSentences(qa.answerText);
  const scored = sentences
    .map((s) => ({ s, score: diceSimilarity(question, s) }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 3);

  const best = scored[0];
  if (!best || best.score < 0.2) {
    return {
      mode: 'offline',
      text: '我目前处于免费离线模式：我会优先从你当前答案里找最相关的句子。这个问题我没找到明显匹配的内容。你可以：\n1) 具体指出你疑惑的是哪一句\n2) 或者开启本地 Ollama（下方勾选并填写模型）',
    };
  }

  const lines = scored.map((x, i) => `${i + 1}. ${x.s}（相似度 ${x.score.toFixed(2)}）`).join('\n');
  return { mode: 'offline', text: `我从当前答案里找到可能相关的句子：\n${lines}` };
}

async function ollamaChat(prompt) {
  const url = (state.settings.ollamaUrl || 'http://localhost:11434').replace(/\/$/, '');
  const model = state.settings.ollamaModel || 'qwen2.5:3b';
  const payload = {
    model,
    messages: [
      { role: 'system', content: '你是一个背诵辅导助手。回答要简洁，优先解释用户当前背诵材料中的内容。' },
      { role: 'user', content: prompt },
    ],
    stream: false,
  };

  const res = await fetch(`${url}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error('ollama_failed');
  const json = await res.json();
  const text = json?.message?.content;
  if (!text) throw new Error('ollama_empty');
  return text;
}

async function ask() {
  const qa = getCurrentQa();
  if (!qa) return;
  const q = (ui.inputAsk.value || '').trim();
  if (!q) return;

  const context = `问题：${qa.question}\n答案：${qa.answerText}\n用户提问：${q}`;

  ui.askAnswer.textContent = '思考中…';

  if (state.settings.useOllama) {
    try {
      const text = await ollamaChat(context);
      ui.askAnswer.textContent = text;
      return;
    } catch {
      const off = offlineAnswer(qa, q);
      ui.askAnswer.textContent = off.text;
      return;
    }
  }

  const off = offlineAnswer(qa, q);
  ui.askAnswer.textContent = off.text;
}

function exportData() {
  const payload = {
    exportedAt: nowIso(),
    version: 1,
    data: state,
  };
  download(`recite-export-${Date.now()}.json`, JSON.stringify(payload, null, 2));
}

async function exportDocx() {
  if (!window.docx) {
    alert('DOCX 导出不可用：docx 库未加载。请刷新页面后重试。');
    return;
  }

  try {
    const { Document, Packer, Paragraph, HeadingLevel, TextRun, HighlightColor } = window.docx;

    const htmlParagraphsToDocxParas = (answerHtml, fallbackText) => {
      const out = [];
      const fallback = String(fallbackText || '').replace(/\r\n/g, '\n').trim();

    // Prefer rich HTML whenever possible (to preserve paragraphs/bold/highlight)
    const rawHtml = sanitizeAnswerHtml(normalizeImportedHtml(answerHtml || ''));
    const htmlText = String(stripHtmlToText(rawHtml || '') || '').trim();

    // Only fall back to plain text when HTML is effectively empty.
    if (!htmlText) {
      if (!fallback) return [new Paragraph({ text: '' })];
      const paras = fallback
        .split(/\n+/g)
        .map((p) => p.trim())
        .filter(Boolean);
      return paras.length
        ? paras.map((p) => new Paragraph({ children: [new TextRun({ text: p })] }))
        : [new Paragraph({ text: '' })];
    }

      const mapColor = (c) => {
        const v = String(c || '').toLowerCase();
        if (v === 'green') return HighlightColor?.GREEN || 'green';
        if (v === 'cyan') return HighlightColor?.CYAN || 'cyan';
        if (v === 'magenta') return HighlightColor?.MAGENTA || 'magenta';
        if (v === 'yellow') return HighlightColor?.YELLOW || 'yellow';
        return HighlightColor?.YELLOW || 'yellow';
      };

      const makeTextRun = (text, style) => {
        const opts = { text: String(text || '') };
        if (style?.bold) opts.bold = true;
        if (style?.highlight) opts.highlight = mapColor(style.highlightColor);
        return new TextRun(opts);
      };

    // Convert a node to multiple paragraphs; split on <br> to avoid collapsing into one big paragraph.
    const nodeToParagraphRunsList = (node, ctx) => {
      const paras = [];
      let cur = [];
      let curHasText = false;

      const pushCur = () => {
        paras.push(curHasText ? cur : [new TextRun({ text: '' })]);
        cur = [];
        curHasText = false;
      };

      const walk = (n, style) => {
        if (!n) return;
        if (n.nodeType === Node.TEXT_NODE) {
          const t = String(n.nodeValue || '');
          if (t) {
            if (t.trim()) curHasText = true;
            cur.push(makeTextRun(t, style));
          }
          return;
        }
        if (n.nodeType !== Node.ELEMENT_NODE) return;
        const tag = (n.tagName || '').toUpperCase();
        if (tag === 'BR') {
          pushCur();
          return;
        }
        const next = { ...style };
        if (tag === 'STRONG' || tag === 'B') next.bold = true;
        if (tag === 'MARK') {
          next.highlight = true;
          const c = String(n.getAttribute?.('data-color') || '').toLowerCase();
          next.highlightColor = c || next.highlightColor;
        }
        Array.from(n.childNodes || []).forEach((c) => walk(c, next));
      };

      walk(node, ctx || { bold: false, highlight: false });
      // flush remaining
      if (cur.length || curHasText) pushCur();
      return paras;
    };

    // Only use TOP-LEVEL blocks to avoid duplicating nested <div>/<p> and collapsing formatting.
    const doc = new DOMParser().parseFromString(rawHtml || '', 'text/html');
    const body = doc.body;
    if (!body || !body.childNodes.length) {
      return [new Paragraph({ text: '' })];
    }

    const blocks = Array.from(body.children || []).filter((n) => {
      const t = (n.tagName || '').toUpperCase();
      return t === 'P' || t === 'DIV';
    });
    const topBlocks = blocks.length ? blocks : [body];

    topBlocks.forEach((b, idx) => {
      const runsList = nodeToParagraphRunsList(b, { bold: false, highlight: false });
      runsList.forEach((runs) => out.push(new Paragraph({ children: runs })));
      // Keep an empty line between top-level blocks (similar to blank line between paragraphs)
      if (idx !== topBlocks.length - 1) out.push(new Paragraph({ text: '' }));
    });

    const anyText = htmlText.length > 0;

    if (!anyText && fallback) {
      const paras = fallback
        .split(/\n+/g)
        .map((p) => p.trim())
        .filter(Boolean);
      return paras.length
        ? paras.map((p) => new Paragraph({ children: [new TextRun({ text: p })] }))
        : [new Paragraph({ text: '' })];
    }

      return out.length ? out : [new Paragraph({ text: '' })];
    };

    const children = [];

    ensureCollections();
    const qas = Array.isArray(state.qas) ? state.qas : [];
    const exportColId = listState.collectionId || state.progress?.currentCollectionId || null;
    if (!exportColId) {
      alert('请先进入某个合集（点进合集后再导出）。');
      return;
    }

    const col = (Array.isArray(state.collections) ? state.collections : []).find((c) => c.id === exportColId);
    if (!col) {
      alert('当前合集不存在或已被删除，请刷新后重试。');
      return;
    }

    const cols = [col];
    cols.forEach((col, ci) => {
      const colName = String(col?.name || '（未命名合集）');
      children.push(
        new Paragraph({
          text: colName,
          heading: HeadingLevel.HEADING_1,
        })
      );

      const inCol = qas.filter((q) => (q.collectionId || DEFAULT_COLLECTION_ID) === col.id);
      inCol.forEach((qa, qi) => {
        const title = (qa.question || '').trim() || '（无标题）';
        children.push(
          new Paragraph({
            text: title,
            heading: HeadingLevel.HEADING_2,
          })
        );

        const paras = htmlParagraphsToDocxParas(qa.answerHtml, qa.answerText);
        paras.forEach((p) => children.push(p));

        const fp = String(qa.focusPoints || '').trim();
        if (fp) {
          children.push(
            new Paragraph({
              text: '背诵需关注的点',
              heading: HeadingLevel.HEADING_3,
            })
          );
          fp
            .split(/\n+/g)
            .map((x) => x.trim())
            .filter(Boolean)
            .forEach((line) => {
              children.push(new Paragraph({ children: [new TextRun({ text: line })] }));
            });
        }

        if (qi !== inCol.length - 1) children.push(new Paragraph({ text: '' }));
      });

      if (ci !== cols.length - 1) children.push(new Paragraph({ text: '' }));
    });

    const doc = new Document({
      sections: [
        {
          properties: {},
          children,
        },
      ],
    });

  const blob = await Packer.toBlob(doc);
  const name = `${String(col?.name || 'recite').replace(/[\\/:*?"<>|]/g, '_')}.docx`;
  download(name, blob);
  } catch (e) {
    console.error(e);
    alert(`DOCX 导出失败：${e?.message || e}`);
  }
}

function importData(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || ''));
      const data = parsed?.data;
      if (!data || !Array.isArray(data.qas)) throw new Error('bad');
      ensureCollections();

      const col = createCollection(normalizeImportedCollectionName(file?.name));
      const importedQas = data.qas
        .filter(Boolean)
        .map((q) => ({
          id: uid(),
          question: q.question || '',
          answerText: q.answerText || '',
          answerHtml: q.answerHtml || '',
          focusPoints: q.focusPoints || '',
          createdAt: nowIso(),
          updatedAt: nowIso(),
          collectionId: col.id,
        }))
        .filter((q) => (q.question || '').trim() || (q.answerText || '').trim());

      if (!importedQas.length) throw new Error('bad');

      state.qas = [...importedQas, ...state.qas];
      state.progress.currentQaId = importedQas[0].id;
      setActiveCollection(col.id);
      saveState();
      stopPlayer();
      render();
      fillEditorFromCurrent();
      updateMatches();
    } catch {
      alert('导入失败：文件格式不正确');
    }
  };
  reader.readAsText(file);
}

ui.btnSave.addEventListener('click', () => {
  flushFocusPoints();
  const cur = getCurrentQa();
  const question = (ui.inputQuestion.value || '').trim();
  const answerHtml = ui.inputAnswerRich ? extractAnswerHtmlFromRichEditor() : '';
  const answerText = (ui.inputAnswerRich ? getPlainAnswerFromRichEditor() : (ui.inputAnswer.value || '')).trim();

  if (!question && !answerText) return;

  const cid = listState.collectionId || cur?.collectionId || DEFAULT_COLLECTION_ID;
  const focusPoints = ui.inputFocusPoints ? String(ui.inputFocusPoints.value || '') : (cur?.focusPoints || '');
  const qa = cur
    ? { ...cur, question, answerText, answerHtml, focusPoints, collectionId: cid, updatedAt: nowIso() }
    : { id: uid(), question, answerText, answerHtml, focusPoints, collectionId: cid, createdAt: nowIso(), updatedAt: nowIso() };

  upsertQa(qa);
  state.progress.currentQaId = qa.id;
  saveState();
  render();
  fillEditorFromCurrent();
});

ui.btnNew.addEventListener('click', () => {
  flushFocusPoints();
  state.progress.currentQaId = null;
  ui.inputQuestion.value = '';
  ui.inputAnswer.value = '';
  if (ui.inputAnswerRich) ui.inputAnswerRich.innerHTML = '';
  if (ui.inputFocusPoints) ui.inputFocusPoints.value = '';
  resetReciteCheck();
  resetCardCheck();
  saveState();
  render();
  fillEditorFromCurrent();
  updateMatches();
});

ui.btnDelete.addEventListener('click', () => {
  flushFocusPoints();
  const qa = getCurrentQa();
  if (!qa) return;
  const ok = confirm('确认删除当前 QA？');
  if (!ok) return;
  stopPlayer();
  deleteCurrentQa();
  listState.selected.delete(qa.id);
  saveState();
  render();
  fillEditorFromCurrent();
  updateMatches();
});

ui.btnApplySettings.addEventListener('click', () => {
  applySettingsFromUI();
  render();
  updateMatches();
});

ui.checkQaReciteSideBySide?.addEventListener('change', () => {
  state.settings.qaReciteSideBySide = !!ui.checkQaReciteSideBySide.checked;
  saveState();
  render();
});

ui.checkReciteOnlyHighlights?.addEventListener('change', () => {
  state.settings.reciteOnlyHighlights = !!ui.checkReciteOnlyHighlights.checked;
  saveState();
  resetReciteCheck();
  render();
  updateMatches();
});

ui.inputQaSearch.addEventListener('input', () => {
  listState.query = ui.inputQaSearch.value || '';
  listState.page = 1;
  render();
});

ui.btnPagePrev.addEventListener('click', () => {
  listState.page = Math.max(1, listState.page - 1);
  render();
});

ui.btnPageNext.addEventListener('click', () => {
  const { totalPages } = getPagedQas();
  listState.page = Math.min(totalPages, listState.page + 1);
  render();
});

ui.checkSelectAll.addEventListener('change', () => {
  const { items } = getPagedQas();
  if (ui.checkSelectAll.checked) {
    items.forEach((x) => listState.selected.add(x.id));
  } else {
    items.forEach((x) => listState.selected.delete(x.id));
  }
  renderQaList();
});

ui.btnDeleteSelected.addEventListener('click', () => deleteSelectedQas());

ui.btnDeleteCollectionAll?.addEventListener('click', () => deleteCurrentCollectionQas());

ui.btnClearSelection.addEventListener('click', () => {
  listState.selected.clear();
  renderQaList();
});

ui.btnDeleteAll.addEventListener('click', () => {
  deleteAllQas();
});

ui.btnTtsTest.addEventListener('click', async () => {
  const v = voicesCount();
  const enabled = !!state.settings.ttsEnabled;
  const ua = navigator.userAgent || '';
  const isAndroid = /Android/i.test(ua);
  ui.statusLine.textContent = `TTS测试：enabled=${enabled} voices=${v} android=${isAndroid}（如果voices=0通常需要安装系统语音包）`;
  await speak('这是朗读测试。');
});

ui.btnStart.addEventListener('click', () => {
  if (player.running && player.paused) return resumePlayer();
  if (player.running) return;
  startPlayer();
});

ui.btnPause.addEventListener('click', () => pausePlayer());
ui.btnStop.addEventListener('click', () => stopPlayer());
ui.btnPrev.addEventListener('click', () => gotoPrevQa());
ui.btnNext.addEventListener('click', () => gotoNextQa());

ui.btnMarkLevel0?.addEventListener('click', () => markCurrentQaToReciteLevel(0));
ui.btnMarkLevel1?.addEventListener('click', () => markCurrentQaToReciteLevel(1));
ui.btnMarkLevel2?.addEventListener('click', () => markCurrentQaToReciteLevel(2));

ui.btnRecStart.addEventListener('click', () => startRecording());
ui.btnRecStop.addEventListener('click', () => stopRecording());
let _autoManualCheckTimer = null;
function scheduleAutoManualCheck() {
  if (_autoManualCheckTimer) clearTimeout(_autoManualCheckTimer);
  _autoManualCheckTimer = setTimeout(() => {
    _autoManualCheckTimer = null;
    applyManualCheck(true);
  }, 300);
}

ui.btnCheck.addEventListener('click', () => {
  reciteCheck.manualAuto = !reciteCheck.manualAuto;
  updateManualCheckButton();
  if (reciteCheck.manualAuto) {
    applyManualCheck(true);
  } else {
    resetReciteCheck();
    ui.speechHint.textContent = '已关闭自动输入检测。';
    updateMatches();
  }
});

ui.btnRepairMaskPlaceholder?.addEventListener('click', () => {
  stopPlayer();
  repairMaskPlaceholderForCurrentQa();
});

ui.inputRecited.addEventListener('input', () => {
  if (reciteCheck.manualAuto) scheduleAutoManualCheck();
  else updateMatches();
});

ui.btnCardMode.addEventListener('click', () => {
  cardCheck.enabled = !cardCheck.enabled;
  resetCardCheck();
  render();
});

ui.btnCardPrev.addEventListener('click', () => {
  cardCheck.index = Math.max(0, cardCheck.index - 1);
  cardCheck.flipped = false;
  renderCurrentQaView();
});

ui.btnCardNext.addEventListener('click', () => {
  const qa = getCurrentQa();
  const total = qa ? splitSentences(qa.answerText).length : 0;
  cardCheck.index = Math.min(Math.max(0, total - 1), cardCheck.index + 1);
  cardCheck.flipped = false;
  renderCurrentQaView();
});

ui.btnCardFlip.addEventListener('click', () => {
  cardCheck.flipped = !cardCheck.flipped;
  renderCurrentQaView();
});

ui.btnMaskToggleAll.addEventListener('click', () => {
  maskState.showAll = !maskState.showAll;
  render();
});

// Floating answer highlight palette
ui.ansMarkPalette?.addEventListener('mousedown', (e) => {
  // Avoid selection collapse before click is handled
  e.preventDefault();
});

ui.ansMarkPalette?.addEventListener('click', (e) => {
  const t = e.target;
  if (!t || !t.getAttribute) return;
  const action = t.getAttribute('data-action');
  const color = t.getAttribute('data-color');
  if (action === 'clear') {
    clearAnswerHighlights();
    hideAnsMarkPalette();
    return;
  }
  if (color) {
    applyAnswerHighlight(color);
    hideAnsMarkPalette();
  }
});

document.addEventListener('selectionchange', () => {
  // Delay to allow selection to settle
  setTimeout(() => showAnsMarkPaletteNearSelection(), 0);
});

document.addEventListener('mousedown', (e) => {
  if (!ui.ansMarkPalette) return;
  if (ui.ansMarkPalette.style.display === 'none') return;
  const t = e.target;
  if (t && ui.ansMarkPalette.contains(t)) return;
  hideAnsMarkPalette();
});

window.addEventListener('scroll', () => {
  // Keep palette near selection when scrolling
  if (!ui.ansMarkPalette || ui.ansMarkPalette.style.display === 'none') return;
  showAnsMarkPaletteNearSelection();
}, true);

window.addEventListener('resize', () => {
  if (!ui.ansMarkPalette || ui.ansMarkPalette.style.display === 'none') return;
  showAnsMarkPaletteNearSelection();
});

ui.btnAsk.addEventListener('click', () => ask());
ui.inputAsk.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') ask();
});

ui.btnExport.addEventListener('click', () => exportData());
ui.btnExportDocx.addEventListener('click', () => exportDocx());
ui.btnQaTimerPause?.addEventListener('click', () => toggleQaTimerPause());
ui.fileImport.addEventListener('change', (e) => {
  const f = e.target.files?.[0];
  if (!f) return;
  const isDocx = f.name.toLowerCase().endsWith('.docx');
  if (isDocx) importDocx(f);
  else importData(f);
  e.target.value = '';
});

// Keyboard shortcuts (PC)
document.addEventListener('keydown', (e) => {
  if (e.defaultPrevented) return;
  if (isTypingTarget(e.target)) return;

  // Save
  if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 's') {
    e.preventDefault();
    ui.btnSave?.click();
    return;
  }

  // Stop
  if (e.key === 'Escape') {
    e.preventDefault();
    stopPlayer();
    return;
  }

  // Start / Pause toggle
  if (e.key === ' ') {
    e.preventDefault();
    if (!player.running) {
      ui.btnStart?.click();
    } else if (player.paused) {
      resumePlayer();
    } else {
      pausePlayer();
    }
    return;
  }

  // Prev/Next QA
  if (e.key === 'ArrowLeft') {
    e.preventDefault();
    gotoPrevQa();
    return;
  }
  if (e.key === 'ArrowRight') {
    e.preventDefault();
    gotoNextQa();
    return;
  }

  // Prev/Next sentence
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    stepSentence(-1);
    return;
  }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    stepSentence(1);
    return;
  }
});

render();
fillEditorFromCurrent();
updateMatches();
startQaTimerForCurrent();
