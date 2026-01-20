(() => {
  // =======================
  // CONFIG & CONSTANTS teste pra kraii
  // =======================
  const TABLE_SELECTOR = "table.po-table";
  const TH_SELECTOR = 'th[data-po-table-column-name]';
  const WRAPPER_SELECTORS = [
    ".po-table-wrapper",
    ".po-table-main-container .po-table-wrapper",
    ".po-table-main-container",
  ];

  const STORAGE_KEY_WIDTHS = "po_col_widths_v6";
  const STORAGE_KEY_FONT = "po_table_font_px_v4";
  const STORAGE_KEY_PANEL_MIN = "po_panel_minimized_v3";
  const STORAGE_KEY_HIDDEN = "po_hidden_cols_v1";
  const STORAGE_KEY_LOCK = "po_layout_lock_v1";
  const STORAGE_KEY_BOXW = "po_box_fixed_w_v1";

  // ARQUIVAMENTO & LOADER (DIÁRIO / LIXEIRA)
  const STORAGE_KEY_ARCHIVED = "po_archived_data_v3";
  const STORAGE_KEY_EXPAND = "po_auto_expand_mode";

  // ✅ tamanho do painel (resizable)
  const STORAGE_KEY_PANEL_SIZE = "po_panel_size_v1";

  // ✅ PAINEL NOTAS (separado)
  const PO_NOTAS_PANEL_ID = "poNotasPanel";

  // ⚠️ Mantém a mesma chave (compatível com versões antigas),
  // mas agora significa: "painel de notas fechado pelo usuário" e deve aparecer o botão de reabrir no painel principal.
  const STORAGE_KEY_NOTAS_PANEL_HIDDEN = "po_notas_panel_hidden_v1";

  const MIN_W = 60,
    MAX_W = 850,
    EXTRA_PAD = 28;
  const FONT_MIN = 10,
    FONT_MAX = 22,
    FONT_DEFAULT = 13;

  // ✅ posição do painel principal
  const PANEL_LEFT = 7,
    PANEL_BOTTOM = 50;

  // ✅ limites de resize do painel principal
  const PANEL_MIN_W = 280;
  const PANEL_MAX_W_MARGIN = 10;
  const PANEL_MIN_H = 180;
  const PANEL_MAX_H_MARGIN = 10;

  const REGEX_DATE = /(\d{2}\/\d{2}\/\d{4})/;

  let _currentSearchTerm = "";

  // ✅ trava global enquanto o usuário está arrastando resize do painel principal
  let __poPanelIsResizing = false;

  // =======================
  // HELPERS
  // =======================
  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }
  function loadJSON(key, fallback) {
    try {
      return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback));
    } catch {
      return fallback;
    }
  }
  function saveJSON(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  function loadWidths() {
    return loadJSON(STORAGE_KEY_WIDTHS, {});
  }
  function saveWidths(map) {
    saveJSON(STORAGE_KEY_WIDTHS, map);
  }
  function loadHidden() {
    const arr = loadJSON(STORAGE_KEY_HIDDEN, []);
    return new Set(Array.isArray(arr) ? arr : []);
  }
  function saveHidden(set) {
    localStorage.setItem(STORAGE_KEY_HIDDEN, JSON.stringify(Array.from(set)));
  }

  function loadArchivedData() {
    return loadJSON(STORAGE_KEY_ARCHIVED, []);
  }
  function saveArchivedData(arr) {
    saveJSON(STORAGE_KEY_ARCHIVED, arr);
  }

  function loadFontPx() {
    const v = Number(localStorage.getItem(STORAGE_KEY_FONT));
    return Number.isFinite(v) ? clamp(v, FONT_MIN, FONT_MAX) : FONT_DEFAULT;
  }
  function saveFontPx(px) {
    localStorage.setItem(STORAGE_KEY_FONT, String(px));
  }
  function loadMinimized() {
    return localStorage.getItem(STORAGE_KEY_PANEL_MIN) === "1";
  }
  function saveMinimized(v) {
    localStorage.setItem(STORAGE_KEY_PANEL_MIN, v ? "1" : "0");
  }
  function loadLock() {
    return localStorage.getItem(STORAGE_KEY_LOCK) === "1";
  }
  function saveLock(v) {
    localStorage.setItem(STORAGE_KEY_LOCK, v ? "1" : "0");
  }
  function loadBoxWidth() {
    const v = Number(localStorage.getItem(STORAGE_KEY_BOXW));
    return Number.isFinite(v) && v > 0 ? v : null;
  }
  function saveBoxWidth(v) {
    if (!v || !Number.isFinite(v) || v <= 0) {
      localStorage.removeItem(STORAGE_KEY_BOXW);
      return;
    }
    localStorage.setItem(STORAGE_KEY_BOXW, String(Math.round(v)));
  }

  const wait = (ms) => new Promise((r) => setTimeout(r, ms));

  // =======================
  // ✅ UTIL: normalização de texto (ÚNICA)
  // =======================
  function _norm(s) {
    return (s || "")
      .toString()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function __poIsElVisible(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    return r.width > 40 && r.height > 40;
  }

  // =======================
  // FONT & LAYOUT
  // =======================
  function applyFont(px) {
    const v = clamp(Math.round(px), FONT_MIN, FONT_MAX);
    document.documentElement.style.setProperty("--po-table-font", `${v}px`);
    saveFontPx(v);

    const label = document.getElementById("poFontValue");
    if (label) label.textContent = `${v}px`;

    const range = document.getElementById("poFontRange");
    if (range && Number(range.value) !== v) range.value = String(v);
  }

  let __poSyncingBox = false;

  function getWrapperEl() {
    for (const sel of WRAPPER_SELECTORS) {
      const el = document.querySelector(sel);
      if (el) return el;
    }
    const table = document.querySelector(TABLE_SELECTOR);
    return table ? table.closest(".po-table-wrapper") || table.parentElement : null;
  }
  function getWrapperFixedWidthNow() {
    const w = getWrapperEl();
    if (!w) return null;
    const v = parseFloat(getComputedStyle(w).width);
    return Number.isFinite(v) ? v : null;
  }
  function setWrapperFixedWidth(px) {
    const w = getWrapperEl();
    if (!w || !px) return;
    const v = `${Math.round(px)}px`;
    w.style.setProperty("width", v, "important");
    w.style.setProperty("min-width", v, "important");
    w.style.setProperty("max-width", v, "important");
  }
  function clearWrapperFixedWidth() {
    const w = getWrapperEl();
    if (!w) return;
    w.style.removeProperty("width");
    w.style.removeProperty("min-width");
    w.style.removeProperty("max-width");
  }

  function sumVisibleHeaderWidthsPreferSaved() {
    const table = document.querySelector(TABLE_SELECTOR);
    if (!table) return null;

    const saved = loadWidths();
    const ths = Array.from(table.querySelectorAll("thead th")).filter((th) => {
      if (th.classList.contains("po-col-hidden")) return false;
      const cs = getComputedStyle(th);
      return !(cs.display === "none" || cs.visibility === "hidden");
    });

    if (!ths.length) return null;

    let total = 0;
    for (const th of ths) {
      const colName = th.getAttribute("data-po-table-column-name");
      const wSaved = colName && Number.isFinite(saved[colName]) ? Number(saved[colName]) : null;
      if (wSaved && wSaved > 0) {
        total += wSaved;
        continue;
      }
      const wNow = th.getBoundingClientRect().width;
      if (Number.isFinite(wNow) && wNow > 0) total += wNow;
    }
    return Math.round(total);
  }

  function syncBoxWidthToColumns() {
    if (!loadLock() || __poSyncingBox) return;
    const table = document.querySelector(TABLE_SELECTOR);
    if (!table) return;

    __poSyncingBox = true;
    try {
      const total = sumVisibleHeaderWidthsPreferSaved();
      if (!total || total <= 0) return;

      const currentSaved = loadBoxWidth();
      const currentWrapper = getWrapperFixedWidthNow();
      const target = Math.round(total);

      const almostEqual = (a, b) => a != null && b != null && Math.abs(a - b) <= 1;
      if (almostEqual(currentSaved, target) && almostEqual(currentWrapper, target)) return;

      saveBoxWidth(target);
      setWrapperFixedWidth(target);
    } finally {
      __poSyncingBox = false;
    }
  }

  function updateLockUI() {
    const badge = document.getElementById("poLockState");
    const btn = document.getElementById("poLockToggle");
    const locked = loadLock();
    if (badge) badge.textContent = locked ? "Colunas Travadas" : "Colunas Full Page";
    if (btn) btn.textContent = locked ? "🔒" : "🔓";
  }

  function applyLockState(force) {
    if (typeof force === "boolean") saveLock(force);
    const locked = loadLock();
    const table = document.querySelector(TABLE_SELECTOR);
    if (!table) {
      updateLockUI();
      return;
    }

    if (!locked) {
      clearWrapperFixedWidth();
      saveBoxWidth(null);
      table.style.removeProperty("table-layout");
      updateLockUI();
      return;
    }

    table.style.setProperty("table-layout", "fixed", "important");
    const savedBox = loadBoxWidth();
    if (savedBox && savedBox > 0) setWrapperFixedWidth(savedBox);
    else syncBoxWidthToColumns();

    updateLockUI();
  }

  // ============================================
  // ✅ PAINEL PRINCIPAL RESIZABLE (inferior esquerdo)
  // ============================================
  function loadPanelSize() {
    const v = loadJSON(STORAGE_KEY_PANEL_SIZE, null);
    if (!v || typeof v !== "object") return { w: null, h: null };
    const w = Number(v.w),
      h = Number(v.h);
    return {
      w: Number.isFinite(w) && w > 0 ? w : null,
      h: Number.isFinite(h) && h > 0 ? h : null,
    };
  }
  function savePanelSize(w, h) {
    const payload = { w: Math.round(w), h: Math.round(h) };
    saveJSON(STORAGE_KEY_PANEL_SIZE, payload);
  }

  function getPanelBottomPx(panel) {
    const cs = getComputedStyle(panel);
    const b = parseFloat(cs.bottom);
    return Number.isFinite(b) ? b : PANEL_BOTTOM;
  }
  function getPanelLeftPx(panel) {
    const cs = getComputedStyle(panel);
    const l = parseFloat(cs.left);
    return Number.isFinite(l) ? l : PANEL_LEFT;
  }

  function clampPanelSizeToViewport(panel, desiredW, desiredH) {
    const bottomPx = getPanelBottomPx(panel);
    const leftPx = getPanelLeftPx(panel);

    const maxW = Math.max(PANEL_MIN_W, window.innerWidth - leftPx - PANEL_MAX_W_MARGIN);
    const maxH = Math.max(PANEL_MIN_H, window.innerHeight - bottomPx - PANEL_MAX_H_MARGIN);

    const w = clamp(desiredW, PANEL_MIN_W, maxW);
    const h = clamp(desiredH, PANEL_MIN_H, maxH);

    return { w, h, maxW, maxH };
  }

  function applyPanelSize(panel, size, opts = {}) {
    if (!panel) return;
    if (panel.classList.contains("minimized")) return;

    const allowSave = opts.allowSave !== false;

    const bottomPx = getPanelBottomPx(panel);
    const leftPx = getPanelLeftPx(panel);

    const maxW = Math.max(PANEL_MIN_W, window.innerWidth - leftPx - PANEL_MAX_W_MARGIN);
    const maxH = Math.max(PANEL_MIN_H, window.innerHeight - bottomPx - PANEL_MAX_H_MARGIN);

    panel.style.setProperty("max-width", `${maxW}px`, "important");
    panel.style.setProperty("max-height", `${maxH}px`, "important");

    const rect = panel.getBoundingClientRect();
    const desiredW = size && size.w ? Number(size.w) : rect.width;
    const desiredH = size && size.h ? Number(size.h) : rect.height;

    const clamped = clampPanelSizeToViewport(panel, desiredW, desiredH);

    panel.style.setProperty("width", `${Math.round(clamped.w)}px`, "important");
    panel.style.setProperty("height", `${Math.round(clamped.h)}px`, "important");

    if (allowSave && !__poPanelIsResizing) {
      savePanelSize(clamped.w, clamped.h);
    }
  }

  function installPanelResizers(panel) {
    if (!panel || panel.dataset.poResizable === "1") return;
    panel.dataset.poResizable = "1";

    const right = document.createElement("div");
    right.className = "po-panel-resize-right";
    right.title = "Ajustar largura (ABC)";

    const top = document.createElement("div");
    top.className = "po-panel-resize-top";
    top.title = "Ajustar altura (↕)";

    panel.appendChild(right);
    panel.appendChild(top);

    const startResize = (kind, ev, handleEl) => {
      ev.preventDefault();
      ev.stopPropagation();
      if (panel.classList.contains("minimized")) return;

      __poPanelIsResizing = true;
      document.body.classList.add("po-panel-resizing");

      handleEl.setPointerCapture?.(ev.pointerId);

      const startX = ev.clientX;
      const startY = ev.clientY;

      const rect = panel.getBoundingClientRect();
      const startW = rect.width;
      const startH = rect.height;

      const onMove = (e) => {
        if (kind === "w") {
          const dx = e.clientX - startX;
          const newW = startW + dx;
          const clamped = clampPanelSizeToViewport(panel, newW, startH);
          panel.style.setProperty("width", `${Math.round(clamped.w)}px`, "important");
        } else {
          const dy = e.clientY - startY;
          const newH = startH - dy;
          const clamped = clampPanelSizeToViewport(panel, startW, newH);
          panel.style.setProperty("height", `${Math.round(clamped.h)}px`, "important");
        }
      };

      const onUp = () => {
        document.removeEventListener("pointermove", onMove, true);
        document.removeEventListener("pointerup", onUp, true);

        document.body.classList.remove("po-panel-resizing");
        __poPanelIsResizing = false;

        const r2 = panel.getBoundingClientRect();
        const clamped = clampPanelSizeToViewport(panel, r2.width, r2.height);

        panel.style.setProperty("width", `${Math.round(clamped.w)}px`, "important");
        panel.style.setProperty("height", `${Math.round(clamped.h)}px`, "important");

        savePanelSize(clamped.w, clamped.h);
        applyPanelSize(panel, { w: clamped.w, h: clamped.h }, { allowSave: false });
      };

      document.addEventListener("pointermove", onMove, true);
      document.addEventListener("pointerup", onUp, true);
    };

    right.addEventListener("pointerdown", (e) => startResize("w", e, right));
    top.addEventListener("pointerdown", (e) => startResize("h", e, top));

    window.addEventListener(
      "resize",
      () => {
        if (!panel || panel.classList.contains("minimized")) return;
        if (__poPanelIsResizing) return;
        applyPanelSize(panel, loadPanelSize());
      },
      { passive: true }
    );
  }

  // ============================================
  // TURBO LOADER (AUTO EXPAND) - LIXEIRA/DIÁRIO
  // ============================================
  async function startSmartAutoExpand() {
    if (localStorage.getItem(STORAGE_KEY_EXPAND) !== "true") return;

    console.log("[SGEasy] Turbo Loader Iniciado...");

    const overlay = document.createElement("div");
    overlay.id = "poAutoExpandOverlay";
    overlay.style.cssText = `
      position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
      background: rgba(255,255,255,0.95); z-index: 999999;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      font-family: sans-serif; color: #333;
    `;
    overlay.innerHTML = `
      <div style="font-size: 20px; font-weight: bold; margin-bottom: 10px;">Restaurando Lista Completa...</div>
      <div id="poLoaderStatus" style="margin-top: 5px; font-weight: bold; color: #e65100;">Iniciando...</div>
      <button id="poCancelExpandBtn" style="margin-top: 20px; padding: 8px 16px; background: #d32f2f; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">
        ❌ Fechar / Cancelar
      </button>
    `;
    document.body.appendChild(overlay);

    const finishExpansion = () => {
      localStorage.removeItem(STORAGE_KEY_EXPAND);
      if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
      runArchivedCleanup();
    };

    document.getElementById("poCancelExpandBtn").addEventListener("click", finishExpansion);
    setTimeout(() => {
      if (document.body.contains(overlay)) finishExpansion();
    }, 20000);

    const statusLabel = document.getElementById("poLoaderStatus");
    await wait(1500);

    let consecutiveMisses = 0;
    let stuckLoadingCount = 0;
    const MAX_MISSES = 5;

    const intervalId = setInterval(() => {
      const footerContainer = document.querySelector(".po-table-footer-show-more");
      let loadMoreBtn = null;

      if (footerContainer) {
        loadMoreBtn = footerContainer.querySelector("button") || footerContainer.querySelector("po-button");
      } else {
        const allBtns = Array.from(document.querySelectorAll("button, po-button"));
        loadMoreBtn = allBtns.find((b) => (b.innerText || "").toLowerCase().includes("carregar mais"));
      }

      if (loadMoreBtn) {
        consecutiveMisses = 0;

        if (
          loadMoreBtn.disabled ||
          loadMoreBtn.hasAttribute("disabled") ||
          getComputedStyle(loadMoreBtn).display === "none"
        ) {
          statusLabel.innerHTML = "Carregando dados...";
          stuckLoadingCount++;

          if (stuckLoadingCount > 10) {
            clearInterval(intervalId);
            statusLabel.style.color = "#43a047";
            statusLabel.innerHTML = "Finalizando...";
            setTimeout(finishExpansion, 200);
          }
        } else {
          stuckLoadingCount = 0;
          statusLabel.innerHTML = "Expandindo lista... ⏬";
          loadMoreBtn.click();
        }
      } else {
        consecutiveMisses++;
        statusLabel.innerHTML = "Finalizando...";

        if (consecutiveMisses >= MAX_MISSES) {
          clearInterval(intervalId);
          statusLabel.style.color = "#43a047";
          statusLabel.innerHTML = "Pronto! ✅";
          setTimeout(finishExpansion, 500);
        }
      }
    }, 200);
  }

  // ============================================
  // ✅ TABELA DIÁRIO (para lixeira) - mantém igual
  // ============================================
  function getArchiveTable() {
    const tables = Array.from(document.querySelectorAll(TABLE_SELECTOR));
    if (!tables.length) return null;

    for (const t of tables) {
      const ths = Array.from(t.querySelectorAll(`thead ${TH_SELECTOR}`));
      const names = ths.map((th) => _norm(th.getAttribute("data-po-table-column-name")));
      const hasDisc = names.includes("disciplina");
      const hasCurso = names.includes("curso");
      const hasTurma =
        names.includes("cod. turma") ||
        names.includes("cód. turma") ||
        names.includes("cod turma") ||
        names.includes("codigo turma");
      if (hasDisc && hasCurso && hasTurma) return t;
    }

    if (tables.length === 1) return tables[0];
    return null;
  }

  // =========================================================
  // ✅ NOTAS — TABELA + PAINEL (separado, inferior direito)
  // =========================================================
  function __poNotasPanelHiddenLoad() {
    return localStorage.getItem(STORAGE_KEY_NOTAS_PANEL_HIDDEN) === "1";
  }
  function __poNotasPanelHiddenSave(v) {
    localStorage.setItem(STORAGE_KEY_NOTAS_PANEL_HIDDEN, v ? "1" : "0");
  }

  function __poInjectNotasStylesOnce() {
    if (document.getElementById("poNotasPanelStyles")) return;
    const st = document.createElement("style");
    st.id = "poNotasPanelStyles";
    st.textContent = `
      #${PO_NOTAS_PANEL_ID}{
        position:fixed; right:10px; bottom:50px; z-index:2147483647;
        width:360px; max-width:min(420px, calc(100vw - 20px));
        background:#1d2a33; color:#fff;
        border:1px solid rgba(255,255,255,0.14);
        border-radius:12px; box-shadow:0 10px 30px rgba(0,0,0,0.35);
        font-family: Arial, sans-serif; overflow:hidden;
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-head{
        display:flex; align-items:center; justify-content:space-between;
        padding:10px; border-bottom:1px solid rgba(255,255,255,0.10);
        background:rgba(255,255,255,0.04);
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-title{ font-weight:800; font-size:13px; letter-spacing:.2px; }
      #${PO_NOTAS_PANEL_ID} .po-notas-close{
        width:30px; height:30px; display:grid; place-items:center;
        border-radius:10px; cursor:pointer;
        border:1px solid rgba(255,255,255,0.14);
        background:rgba(255,255,255,0.08);
        user-select:none; font-weight:900;
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-close:hover{ background:rgba(255,255,255,0.14); }
      #${PO_NOTAS_PANEL_ID} .po-notas-body{ padding:10px; display:flex; flex-direction:column; gap:8px; }
      #${PO_NOTAS_PANEL_ID} .po-notas-row{ display:flex; gap:8px; }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn{
        flex:1; border:1px solid rgba(255,255,255,0.14);
        background:rgba(255,255,255,0.08); color:#fff;
        border-radius:10px; padding:8px 10px; cursor:pointer;
        font-weight:700; font-size:12px;
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn:hover{ background:rgba(255,255,255,0.14); }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn.green{ background:rgba(67,160,71,.18); border-color:rgba(67,160,71,.35); }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn.green:hover{ background:rgba(67,160,71,.26); }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn.orange{ background:rgba(255,152,0,.16); border-color:rgba(255,152,0,.32); }
      #${PO_NOTAS_PANEL_ID} .po-notas-btn.orange:hover{ background:rgba(255,152,0,.24); }
      #${PO_NOTAS_PANEL_ID} .po-notas-drop{
        border:1px dashed rgba(255,255,255,0.28);
        background:rgba(255,255,255,0.06);
        border-radius:12px; padding:10px; text-align:center;
        font-size:12px; opacity:.95;
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-drop.dragover{
        background:rgba(255,255,255,0.12);
        border-color:rgba(255,255,255,0.42);
      }
      #${PO_NOTAS_PANEL_ID} .po-notas-status{ font-size:11px; opacity:.85; line-height:1.25; white-space:pre-wrap; }
    `;
    document.head.appendChild(st);
  }

  // ✅ coleta tabelas no document + iframes (same-origin)
  function __poCollectTablesDeep(selectors = ["table.po-table", "po-table table", "table"]) {
    const out = [];
    const visited = new Set();

    const walk = (doc) => {
      if (!doc || visited.has(doc)) return;
      visited.add(doc);

      for (const sel of selectors) {
        doc.querySelectorAll(sel).forEach((t) => out.push(t));
      }

      const iframes = Array.from(doc.querySelectorAll("iframe"));
      for (const fr of iframes) {
        try {
          if (fr.contentDocument) walk(fr.contentDocument);
        } catch {
          // cross-origin
        }
      }
    };

    walk(document);
    return out;
  }

  function __poGetHeadersFromTable(table) {
    const ths = Array.from(table.querySelectorAll("thead th"));
    if (ths.length) {
      return ths.map((th) => _norm(th.getAttribute("data-po-table-column-name") || th.innerText || ""));
    }

    const firstRowTds = Array.from(table.querySelectorAll("tbody tr:first-child td"));
    if (firstRowTds.length) {
      return firstRowTds.map((td) => _norm(td.innerText || ""));
    }

    return [];
  }

  function __poScoreNotasTable(table) {
    const headers = __poGetHeadersFromTable(table).filter(Boolean);
    const headerText = headers.join(" | ");

    const hasAluno = headers.some(
      (h) =>
        ["aluno", "discente", "estudante", "nome", "aprendiz"].includes(h) ||
        h.includes("aluno") ||
        h.includes("nome")
    );

    const hasMat = headers.some(
      (h) => h.includes("matricula") || h.includes("matrícula") || h === "ra" || h.includes("registro")
    );

    const gradeHints = [
      "nota",
      "notas",
      "media",
      "média",
      "avaliacao",
      "avaliação",
      "prova",
      "trabalho",
      "pontuacao",
      "pontuação",
      "pontos",
      "final",
      "n1",
      "n2",
      "n3",
      "n4",
      "a1",
      "a2",
      "a3",
      "a4",
      "av1",
      "av2",
      "recuperacao",
      "recuperação",
      "rec",
    ];

    const hasGradeHeader = headers.some((h) => gradeHints.includes(h) || gradeHints.some((k) => h.includes(k)));
    const hasShortGradeCols = headers.some((h) => /^n\d+$/.test(h) || /^a\d+$/.test(h) || /^av\d+$/.test(h));

    const rowCount = table.querySelectorAll("tbody tr").length;
    const colCount =
      Math.max(0, table.querySelectorAll("thead th").length) ||
      table.querySelectorAll("tbody tr:first-child td").length;

    let score = 0;
    if (hasAluno) score += 6;
    if (hasMat) score += 3;
    if (hasGradeHeader) score += 8;
    if (hasShortGradeCols) score += 4;

    if (rowCount >= 5) score += 2;
    if (rowCount >= 15) score += 2;
    if (colCount >= 6) score += 2;
    if (colCount >= 10) score += 2;

    const looksLikeDiario =
      headerText.includes("disciplina") &&
      headerText.includes("curso") &&
      (headerText.includes("turma") || headerText.includes("cod turma"));
    if (looksLikeDiario) score -= 12;

    if (!__poIsElVisible(table)) score -= 6;

    return score;
  }

  function __poGetNotasTable() {
    const candidates = __poCollectTablesDeep(["table.po-table", "po-table table", "table"]);
    if (!candidates.length) return null;

    let best = null;
    let bestScore = -999;

    for (const t of candidates) {
      const s = __poScoreNotasTable(t);
      if (s > bestScore) {
        bestScore = s;
        best = t;
      }
    }

    if (best) {
      console.log(
        "[SGEasy] Notas: candidatas =",
        candidates.length,
        "melhor score =",
        bestScore,
        "headers =",
        __poGetHeadersFromTable(best)
      );
    }

    if (!best || bestScore < 8) return null;
    return best;
  }

  function __poGetTableScopeRoot(table) {
    if (!table) return document;
    return table.closest(".po-table-main-container") || table.closest(".po-table-wrapper") || table.parentElement || document;
  }

  function __poFindLoadMoreButtonForTable(table) {
    const root = __poGetTableScopeRoot(table);
    const footer = root.querySelector(".po-table-footer-show-more");
    if (footer) return footer.querySelector("button, po-button");

    const candidates = Array.from(root.querySelectorAll("button, po-button"));
    const btn = candidates.find((b) => _norm(b.innerText || "").includes("carregar mais"));
    return btn || null;
  }

  async function __poExpandAllRowsForNotasTable(table, statusLabel, isCancelledFn) {
    let clicks = 0;
    let misses = 0;

    while (clicks < 220) {
      if (isCancelledFn && isCancelledFn()) return { cancelled: true, clicks };

      const btn = __poFindLoadMoreButtonForTable(table);

      if (!btn) {
        misses++;
        if (statusLabel)
          statusLabel.textContent = "Nenhum 'Carregar mais' encontrado (prosseguindo com o que está na tela)...";
        if (misses >= 4) break;
        await wait(150);
        continue;
      }

      const disabled = !!btn.disabled || btn.hasAttribute("disabled") || getComputedStyle(btn).display === "none";
      if (disabled) {
        if (statusLabel) statusLabel.textContent = "Tudo carregado.";
        break;
      }

      if (statusLabel) statusLabel.textContent = `Carregando mais linhas... (${clicks + 1})`;
      btn.click();
      clicks++;

      await wait(190);
    }

    return { cancelled: false, clicks };
  }

  function __poExtractCellValue(cell) {
    if (!cell) return "";

    const input =
      cell.querySelector("input, textarea") ||
      cell.querySelector("po-input input, po-number input, po-decimal input, po-field-container input") ||
      cell.querySelector(".po-field-container input");

    if (input) return (input.value ?? "").toString().trim();

    const roleTb = cell.querySelector('[role="textbox"]');
    if (roleTb) return (roleTb.textContent || "").replace(/\s+/g, " ").trim();

    const clone = cell.cloneNode(true);
    clone
      .querySelectorAll(".po-col-resizer, .po-col-hide-btn, .po-archive-check-container, .po-archive-check, button")
      .forEach((n) => n.remove());

    return (clone.innerText || clone.textContent || "").replace(/\s+/g, " ").trim();
  }

  function __poGetVisibleColumnIndexes(table) {
    const ths = Array.from(table.querySelectorAll("thead th"));
    const idxs = [];
    ths.forEach((th, i) => {
      const cs = getComputedStyle(th);
      const hiddenByClass = th.classList.contains("po-col-hidden");
      const hiddenByStyle = cs.display === "none" || cs.visibility === "hidden";
      if (!hiddenByClass && !hiddenByStyle) idxs.push(i);
    });
    return idxs;
  }

  function __poTableToMatrixVisible(table) {
    const ths = Array.from(table.querySelectorAll("thead th"));
    const idxs = __poGetVisibleColumnIndexes(table);

    const header = idxs.map((i) => __poExtractCellValue(ths[i]));
    const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) => {
      const tds = Array.from(tr.querySelectorAll("td"));
      return idxs.map((i) => __poExtractCellValue(tds[i]));
    });

    return [header, ...rows];
  }

  // =========================================================
  // ✅ XLSX (EXPORT) — gerador ZIP correto (sem erro de extensão/formato)
  // =========================================================
  const __poTextEncoder = new TextEncoder();

  function __poU8(str) {
    return __poTextEncoder.encode(str || "");
  }

  function __poSafeFilePart(s, maxLen = 40) {
    const t = (s || "")
      .toString()
      .trim()
      .replace(/[\\\/:*?"<>|]+/g, "-")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");
    return (t || "SEM_INFO").slice(0, maxLen);
  }

  function __poPad2(n) {
    return String(n).padStart(2, "0");
  }

  function __poNowStamp() {
    const d = new Date();
    return `${d.getFullYear()}-${__poPad2(d.getMonth() + 1)}-${__poPad2(d.getDate())}_${__poPad2(d.getHours())}${__poPad2(
      d.getMinutes()
    )}`;
  }

  function __poEscapeXml(s) {
    return (s ?? "")
      .toString()
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function __poColName(n) {
    let x = n + 1;
    let name = "";
    while (x > 0) {
      const r = (x - 1) % 26;
      name = String.fromCharCode(65 + r) + name;
      x = Math.floor((x - 1) / 26);
    }
    return name;
  }

  // CRC32
  const __poCrcTable = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[i] = c >>> 0;
    }
    return t;
  })();

  function __poCrc32(u8) {
    let crc = 0xffffffff;
    for (let i = 0; i < u8.length; i++) {
      crc = __poCrcTable[(crc ^ u8[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function __poPushU16(arr, n) {
    arr.push(n & 255, (n >>> 8) & 255);
  }
  function __poPushU32(arr, n) {
    arr.push(n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255);
  }

  function __poConcatU8(chunks) {
    let total = 0;
    for (const c of chunks) total += c.length;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.length;
    }
    return out;
  }

  // ZIP store (método 0) — central directory CORRETO
  function __poZipStore(entries) {
    // entries: [{name: string, data: Uint8Array}]
    const localChunks = [];
    const centralChunks = [];

    let offset = 0;

    for (const ent of entries) {
      const nameU8 = __poU8(ent.name);
      const dataU8 = ent.data instanceof Uint8Array ? ent.data : new Uint8Array(ent.data);
      const crc = __poCrc32(dataU8);

      // local header
      const lh = [];
      // signature
      __poPushU32(lh, 0x04034b50);
      __poPushU16(lh, 20); // version needed
      __poPushU16(lh, 0); // flags
      __poPushU16(lh, 0); // method store
      __poPushU16(lh, 0); // time
      __poPushU16(lh, 0); // date
      __poPushU32(lh, crc);
      __poPushU32(lh, dataU8.length);
      __poPushU32(lh, dataU8.length);
      __poPushU16(lh, nameU8.length);
      __poPushU16(lh, 0); // extra len

      const localHeaderU8 = __poConcatU8([new Uint8Array(lh), nameU8]);
      localChunks.push(localHeaderU8, dataU8);

      // central header
      const ch = [];
      __poPushU32(ch, 0x02014b50); // signature
      __poPushU16(ch, 20); // version made
      __poPushU16(ch, 20); // version needed
      __poPushU16(ch, 0); // flags
      __poPushU16(ch, 0); // method
      __poPushU16(ch, 0); // time
      __poPushU16(ch, 0); // date
      __poPushU32(ch, crc);
      __poPushU32(ch, dataU8.length);
      __poPushU32(ch, dataU8.length);
      __poPushU16(ch, nameU8.length);
      __poPushU16(ch, 0); // extra len
      __poPushU16(ch, 0); // comment len
      __poPushU16(ch, 0); // disk start
      __poPushU16(ch, 0); // internal attrs
      __poPushU32(ch, 0); // external attrs
      __poPushU32(ch, offset); // local header offset

      const centralHeaderU8 = __poConcatU8([new Uint8Array(ch), nameU8]);
      centralChunks.push(centralHeaderU8);

      offset += localHeaderU8.length + dataU8.length;
    }

    const centralStart = offset;
    const centralDir = __poConcatU8(centralChunks);
    offset += centralDir.length;

    // end of central directory
    const eocd = [];
    __poPushU32(eocd, 0x06054b50);
    __poPushU16(eocd, 0); // disk
    __poPushU16(eocd, 0); // start disk
    __poPushU16(eocd, entries.length);
    __poPushU16(eocd, entries.length);
    __poPushU32(eocd, centralDir.length);
    __poPushU32(eocd, centralStart);
    __poPushU16(eocd, 0); // comment len

    return __poConcatU8([...localChunks, centralDir, new Uint8Array(eocd)]);
  }

  function __poIsNumericCell(v) {
    const s = (v ?? "").toString().trim();
    if (!s) return false;
    // aceita 10 / 10,5 / 10.5 / -2
    return /^-?\d+(?:[.,]\d+)?$/.test(s);
  }

  function __poToNumberText(v) {
    return (v ?? "").toString().trim().replace(",", ".");
  }

  function __poMatrixToXlsxBytes(matrix, sheetNameRaw) {
    const sheetName = (sheetNameRaw || "Notas").toString().trim().slice(0, 31).replace(/[\[\]\*\/\\\?\:]/g, "-");

    const rowsXml = [];
    let maxC = 1;

    for (let r = 0; r < matrix.length; r++) {
      const row = matrix[r] || [];
      maxC = Math.max(maxC, row.length || 1);

      const cellsXml = [];
      for (let c = 0; c < row.length; c++) {
        const addr = `${__poColName(c)}${r + 1}`;
        const val = row[c] ?? "";
        const s = (val ?? "").toString();
        if (!s.trim()) continue;

        // cabeçalho sempre texto
        if (r === 0 || !__poIsNumericCell(s)) {
          cellsXml.push(
            `<c r="${addr}" t="inlineStr"><is><t>${__poEscapeXml(s)}</t></is></c>`
          );
        } else {
          cellsXml.push(`<c r="${addr}"><v>${__poEscapeXml(__poToNumberText(s))}</v></c>`);
        }
      }

      if (cellsXml.length) {
        rowsXml.push(`<row r="${r + 1}">${cellsXml.join("")}</row>`);
      }
    }

    const dim = `A1:${__poColName(maxC - 1)}${Math.max(1, matrix.length)}`;

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dim}"/>
  <sheetData>
    ${rowsXml.join("")}
  </sheetData>
</worksheet>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${__poEscapeXml(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

    const nowIso = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");

    const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${__poEscapeXml(sheetName)}</dc:title>
  <dc:creator>SGEasy</dc:creator>
  <cp:lastModifiedBy>SGEasy</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:modified>
</cp:coreProperties>`;

    const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>SGEasy</Application>
</Properties>`;

    const entries = [
      { name: "[Content_Types].xml", data: __poU8(contentTypes) },
      { name: "_rels/.rels", data: __poU8(relsXml) },
      { name: "docProps/core.xml", data: __poU8(coreXml) },
      { name: "docProps/app.xml", data: __poU8(appXml) },
      { name: "xl/workbook.xml", data: __poU8(workbookXml) },
      { name: "xl/_rels/workbook.xml.rels", data: __poU8(workbookRels) },
      { name: "xl/styles.xml", data: __poU8(stylesXml) },
      { name: "xl/worksheets/sheet1.xml", data: __poU8(sheetXml) },
    ];

    return __poZipStore(entries);
  }

  function __poDownloadBytes(u8, filename) {
    const blob = new Blob([u8], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  }

  function __poGetNotasContext(table) {
    const root = __poGetTableScopeRoot(table) || document;

    let text = "";

    try {
      const tRect = table.getBoundingClientRect();
      const candidates = Array.from(
        root.querySelectorAll("h1,h2,h3,h4,po-page-header,po-info,po-field-container,label,span,div")
      );

      const near = [];
      for (const el of candidates) {
        const s = (el.innerText || "").trim();
        if (!s || s.length > 180) continue;
        if (!/turma|disciplina|c[oó]d\.?\s*turma/i.test(s)) continue;

        const r = el.getBoundingClientRect();
        if (r.bottom <= tRect.top + 30 && r.bottom >= -1500) {
          near.push({ s, y: r.bottom });
        }
      }
      near.sort((a, b) => b.y - a.y);
      text = near.slice(0, 10).map((x) => x.s).join("\n");
    } catch {
      text = "";
    }

    const turmaMatch =
      text.match(/c[oó]d\.?\s*turma\s*[:\-]?\s*([A-Za-z0-9._-]+)/i) ||
      text.match(/turma\s*[:\-]?\s*([A-Za-z0-9._-]+)/i);

    const discMatch =
      text.match(/disciplina\s*[:\-]?\s*([^\n\r]{3,80})/i) ||
      text.match(/componente\s*curricular\s*[:\-]?\s*([^\n\r]{3,80})/i);

    const turma = turmaMatch ? (turmaMatch[1] || "").trim() : "";
    let disciplina = discMatch ? (discMatch[1] || "").trim() : "";

    // corta "Curso" ou "Turma" colado
    disciplina = disciplina.replace(/\bcurso\b.*$/i, "").replace(/\bturma\b.*$/i, "").trim();

    return { turma, disciplina };
  }

  function __poMakeNotasFilename(table) {
    const ctx = __poGetNotasContext(table);
    const turmaPart = ctx.turma ? __poSafeFilePart(ctx.turma, 22) : "SEM_TURMA";
    const discPart = ctx.disciplina ? __poSafeFilePart(ctx.disciplina, 28) : "SEM_DISCIPLINA";
    return `Notas_${turmaPart}_${discPart}_${__poNowStamp()}.xlsx`;
  }

  // =========================================================
  // ✅ XLSX (IMPORT) — leitor sem biblioteca (ZIP + XML)
  // =========================================================
  function __poReadU16(view, off) {
    return view.getUint16(off, true);
  }
  function __poReadU32(view, off) {
    return view.getUint32(off, true);
  }

  function __poU8ToStr(u8) {
    return new TextDecoder("utf-8").decode(u8);
  }

  async function __poInflateMaybe(u8) {
    if (typeof DecompressionStream === "undefined") return null;
    const tryFmt = async (fmt) => {
      const ds = new DecompressionStream(fmt);
      const stream = new Blob([u8]).stream().pipeThrough(ds);
      const ab = await new Response(stream).arrayBuffer();
      return new Uint8Array(ab);
    };
    try {
      return await tryFmt("deflate-raw");
    } catch {
      try {
        return await tryFmt("deflate");
      } catch {
        return null;
      }
    }
  }

  async function __poUnzipToMap(arrayBuffer) {
    const u8 = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);

    // procura EOCD no final
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error("ZIP inválido (EOCD não encontrado).");

    const cdSize = __poReadU32(view, eocd + 12);
    const cdOff = __poReadU32(view, eocd + 16);

    let p = cdOff;
    const out = new Map();

    while (p < cdOff + cdSize) {
      const sig = __poReadU32(view, p);
      if (sig !== 0x02014b50) break;

      const method = __poReadU16(view, p + 10);
      const compSize = __poReadU32(view, p + 20);
      const uncompSize = __poReadU32(view, p + 24);
      const nameLen = __poReadU16(view, p + 28);
      const extraLen = __poReadU16(view, p + 30);
      const commentLen = __poReadU16(view, p + 32);
      const localOff = __poReadU32(view, p + 42);

      const name = __poU8ToStr(u8.slice(p + 46, p + 46 + nameLen));

      // local header
      const lSig = __poReadU32(view, localOff);
      if (lSig !== 0x04034b50) {
        p += 46 + nameLen + extraLen + commentLen;
        continue;
      }
      const lNameLen = __poReadU16(view, localOff + 26);
      const lExtraLen = __poReadU16(view, localOff + 28);
      const dataOff = localOff + 30 + lNameLen + lExtraLen;

      const comp = u8.slice(dataOff, dataOff + compSize);

      let data = null;
      if (method === 0) {
        data = comp;
      } else if (method === 8) {
        const inflated = await __poInflateMaybe(comp);
        if (!inflated) throw new Error("Não consegui descompactar XLSX (deflate).");
        data = inflated;
      } else {
        throw new Error("Método de compressão ZIP não suportado: " + method);
      }

      // sem validação de tamanho (uncompSize) para tolerância
      out.set(name, data);

      p += 46 + nameLen + extraLen + commentLen;
    }

    return out;
  }

  function __poParseXml(text) {
    return new DOMParser().parseFromString(text, "application/xml");
  }

  function __poReadSharedStrings(xmlText) {
    const doc = __poParseXml(xmlText);
    const sis = Array.from(doc.getElementsByTagName("si"));
    const out = [];
    for (const si of sis) {
      const ts = Array.from(si.getElementsByTagName("t"));
      let s = "";
      for (const t of ts) s += t.textContent || "";
      out.push(s);
    }
    return out;
  }

  function __poColIndexFromRef(ref) {
    const m = (ref || "").match(/^([A-Z]+)\d+$/i);
    if (!m) return 0;
    const letters = m[1].toUpperCase();
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n - 1; // zero-based
  }

  function __poReadSheetMatrix(sheetXmlText, sharedStrings) {
    const doc = __poParseXml(sheetXmlText);

    const rows = Array.from(doc.getElementsByTagName("row"));
    const rowMap = new Map();

    let maxRow = 0;
    let maxCol = 0;

    for (const rowEl of rows) {
      const rAttr = rowEl.getAttribute("r");
      const rNum = rAttr ? parseInt(rAttr, 10) : null;
      const r = Number.isFinite(rNum) ? rNum : (maxRow + 1);
      maxRow = Math.max(maxRow, r);

      const cells = Array.from(rowEl.getElementsByTagName("c"));
      const line = rowMap.get(r) || [];

      for (const cEl of cells) {
        const ref = cEl.getAttribute("r") || "";
        const c = __poColIndexFromRef(ref);
        maxCol = Math.max(maxCol, c + 1);

        const t = cEl.getAttribute("t") || "";
        let value = "";

        if (t === "s") {
          const vEl = cEl.getElementsByTagName("v")[0];
          const idx = vEl ? parseInt(vEl.textContent || "0", 10) : 0;
          value = (sharedStrings && sharedStrings[idx]) || "";
        } else if (t === "inlineStr") {
          const isEl = cEl.getElementsByTagName("is")[0];
          if (isEl) {
            const ts = Array.from(isEl.getElementsByTagName("t"));
            value = ts.map((x) => x.textContent || "").join("");
          } else {
            value = "";
          }
        } else {
          const vEl = cEl.getElementsByTagName("v")[0];
          value = vEl ? (vEl.textContent || "") : "";
        }

        while (line.length < c + 1) line.push("");
        line[c] = value;
      }

      rowMap.set(r, line);
    }

    // monta matrix preenchendo vazios
    const matrix = [];
    const rowsCount = Math.max(1, maxRow);
    const colsCount = Math.max(1, maxCol);

    for (let r = 1; r <= rowsCount; r++) {
      const line = rowMap.get(r) || [];
      const out = new Array(colsCount).fill("");
      for (let c = 0; c < colsCount; c++) out[c] = line[c] ?? "";
      // descarta linhas totalmente vazias no final
      matrix.push(out);
    }

    // remove linhas vazias finais
    while (matrix.length > 1 && matrix[matrix.length - 1].every((x) => !(x || "").toString().trim())) {
      matrix.pop();
    }

    // remove colunas vazias finais (com segurança)
    let cut = colsCount;
    for (let c = colsCount - 1; c >= 0; c--) {
      const any = matrix.some((row) => (row[c] || "").toString().trim() !== "");
      if (any) break;
      cut = c;
    }
    if (cut < colsCount) {
      for (let r = 0; r < matrix.length; r++) matrix[r] = matrix[r].slice(0, cut);
    }

    return matrix;
  }

  async function __poReadXlsxArrayBuffer(arrayBuffer) {
    const map = await __poUnzipToMap(arrayBuffer);

    const wb = map.get("xl/workbook.xml");
    const rels = map.get("xl/_rels/workbook.xml.rels");
    if (!wb || !rels) throw new Error("XLSX inválido: workbook.xml/rels ausente.");

    const wbDoc = __poParseXml(__poU8ToStr(wb));
    const sheet = wbDoc.getElementsByTagName("sheet")[0];
    if (!sheet) throw new Error("XLSX inválido: sem planilhas.");

    const rid = sheet.getAttribute("r:id") || sheet.getAttribute("id") || "rId1";

    const relDoc = __poParseXml(__poU8ToStr(rels));
    const relNodes = Array.from(relDoc.getElementsByTagName("Relationship"));
    const rel = relNodes.find((x) => (x.getAttribute("Id") || "") === rid);

    let target = rel ? rel.getAttribute("Target") || "" : "worksheets/sheet1.xml";
    if (!target) target = "worksheets/sheet1.xml";
    if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\/+/, "");

    const sheetU8 = map.get(target);
    if (!sheetU8) throw new Error("XLSX inválido: sheet não encontrada (" + target + ").");

    const sharedU8 = map.get("xl/sharedStrings.xml");
    const sharedStrings = sharedU8 ? __poReadSharedStrings(__poU8ToStr(sharedU8)) : null;

    const matrix = __poReadSheetMatrix(__poU8ToStr(sheetU8), sharedStrings);
    const headers = (matrix[0] || []).map((x) => (x ?? "").toString().trim());
    const data = matrix
      .slice(1)
      .filter((r) => Array.isArray(r) && r.some((v) => (v ?? "").toString().trim() !== ""))
      .map((r) => r.map((v) => (v ?? "").toString()));

    return { headers, data };
  }

  function __poGuessCsvDelimiter(text) {
    const lines = (text || "").split(/\r?\n/).slice(0, 10);
    const counts = {
      ";": 0,
      ",": 0,
      "\t": 0,
    };
    for (const line of lines) {
      counts[";"] += (line.match(/;/g) || []).length;
      counts[","] += (line.match(/,/g) || []).length;
      counts["\t"] += (line.match(/\t/g) || []).length;
    }
    let best = ";";
    let bestV = -1;
    for (const k of Object.keys(counts)) {
      if (counts[k] > bestV) {
        bestV = counts[k];
        best = k;
      }
    }
    return best;
  }

  // -------------------------
  // ✅ EXPORT NOTAS (xlsx)
  // -------------------------
  async function __poExportNotasExcel(setStatusFn) {
    const table = __poGetNotasTable();
    if (!table) {
      const msg =
        "Tabela de NOTAS não encontrada nesta tela.\n\n" +
        "Dica: Abra o console (F12) e procure por logs '[SGEasy] Notas:' — isso mostra quantas tabelas ele está enxergando e o cabeçalho da melhor candidata.";
      if (setStatusFn) setStatusFn(msg);
      else alert(msg);
      return;
    }

    let cancelled = false;

    const overlay = document.createElement("div");
    overlay.style.cssText = `
      position:fixed; inset:0; z-index:2147483647;
      background:rgba(255,255,255,.92);
      display:flex; align-items:center; justify-content:center;
      font-family:Arial, sans-serif; color:#222;
    `;
    overlay.innerHTML = `
      <div style="width:min(520px, 92vw); background:#fff; border:1px solid #ddd; border-radius:10px; padding:16px; box-shadow:0 10px 35px rgba(0,0,0,.18)">
        <div style="font-weight:800; font-size:16px;">Exportando Excel da Tabela de Notas...</div>
        <div id="poNotasExportStatus" style="margin-top:8px; font-weight:700; color:#045b8f">Preparando...</div>
        <div style="margin-top:14px; display:flex; justify-content:flex-end; gap:8px;">
          <button id="poNotasExportCancel" style="padding:8px 12px; border:none; border-radius:8px; background:#c62828; color:#fff; cursor:pointer; font-weight:800;">
            Cancelar
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const statusLabel = overlay.querySelector("#poNotasExportStatus");
    overlay.querySelector("#poNotasExportCancel").addEventListener("click", () => {
      cancelled = true;
    });

    try {
      statusLabel.textContent = "Carregando linhas (se houver 'Carregar mais')...";
      const r = await __poExpandAllRowsForNotasTable(table, statusLabel, () => cancelled);
      if (r.cancelled) {
        statusLabel.textContent = "Cancelado.";
        await wait(200);
        return;
      }

      statusLabel.textContent = "Lendo dados da tabela...";
      await wait(80);

      const matrix = __poTableToMatrixVisible(table);
      if (!matrix || matrix.length <= 1) {
        const msg = "Não consegui ler linhas da tabela de notas (talvez a tela ainda esteja carregando).";
        if (setStatusFn) setStatusFn(msg);
        else alert(msg);
        return;
      }

      statusLabel.textContent = `Gerando XLSX (${matrix.length - 1} linhas)...`;
      await wait(80);

      const xlsx = __poMatrixToXlsxBytes(matrix, "Notas");
      const filename = __poMakeNotasFilename(table);
      __poDownloadBytes(xlsx, filename);

      statusLabel.textContent = "Pronto! ✅";
      await wait(250);
    } finally {
      overlay.remove();
    }
  }

  // -------------------------
  // ✅ IMPORT NOTAS (upload)
  // -------------------------
  async function __poReadNotasFile(file) {
    const name = (file?.name || "").toLowerCase();

    // XLSX / XLSM / XLTX
    if (name.endsWith(".xlsx") || name.endsWith(".xlsm") || name.endsWith(".xltx")) {
      // 1) se a lib XLSX existir, usa (mais completa)
      if (window.XLSX) {
        const ab = await file.arrayBuffer();
        const wb = window.XLSX.read(ab, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        const headers = (aoa[0] || []).map((x) => (x ?? "").toString().trim());
        const data = aoa
          .slice(1)
          .filter((r) => Array.isArray(r) && r.some((v) => (v ?? "").toString().trim() !== ""))
          .map((r) => r.map((v) => (v ?? "").toString()));
        return { headers, data };
      }

      // 2) fallback: leitor ZIP+XML (funciona com nosso export e com Excel normal)
      const ab = await file.arrayBuffer();
      return await __poReadXlsxArrayBuffer(ab);
    }

    // CSV (opcional)
    if (name.endsWith(".csv") || name.endsWith(".txt")) {
      const text = await file.text();
      const delim = __poGuessCsvDelimiter(text);
      const rows = text
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean)
        .map((l) => l.split(delim).map((x) => x.replace(/^"|"$/g, "").trim()));

      const headers = (rows[0] || []).map((x) => (x ?? "").toString().trim());
      const data = rows.slice(1).filter((r) => r.some((v) => (v ?? "").toString().trim() !== ""));
      return { headers, data };
    }

    // XLS antigo (HTML) — aceita ainda, se existir
    if (name.endsWith(".xls") || name.endsWith(".html") || name.endsWith(".htm")) {
      const text = await file.text();
      const doc = new DOMParser().parseFromString(text, "text/html");
      const table = doc.querySelector("table");
      if (!table) return null;

      const headers = Array.from(table.querySelectorAll("tr:first-child td, tr:first-child th")).map((c) =>
        (c.textContent || "").trim()
      );

      const rows = Array.from(table.querySelectorAll("tr"))
        .slice(1)
        .map((tr) => Array.from(tr.querySelectorAll("td")).map((td) => (td.textContent || "").trim()));

      const data = rows.filter((r) => r.some((v) => (v || "").trim() !== ""));
      return { headers, data };
    }

    return null;
  }

  function __poGetTableSchema(table) {
    const ths = Array.from(table.querySelectorAll("thead th"));
    if (!ths.length) return null;

    return ths.map((th, idx) => {
      const key = th.getAttribute("data-po-table-column-name") || "";
      const labelRaw = (th.innerText || "").replace(/\s+/g, " ").trim();
      const label = _norm(labelRaw);
      return { idx, th, key: _norm(key), labelRaw, label };
    });
  }

  function __poFindKeyColumnIndex(headersNorm) {
    const candidates = ["matricula", "matrícula", "ra", "registro", "aluno", "nome", "discente"];
    for (let i = 0; i < headersNorm.length; i++) {
      const h = headersNorm[i];
      if (!h) continue;
      if (candidates.includes(h) || candidates.some((k) => h.includes(k))) return i;
    }
    return 0;
  }

  function __poFindKeyColumnIndexInSchema(schema) {
    const candidates = ["matricula", "matrícula", "ra", "registro", "aluno", "nome", "discente"];
    const idx = schema.findIndex((c) => candidates.includes(c.label) || candidates.some((k) => (c.label || "").includes(k)));
    return idx >= 0 ? idx : 0;
  }

  function __poBuildRowMapByKey(table, schema, keySchemaIndex) {
    const map = new Map();
    const rows = Array.from(table.querySelectorAll("tbody tr"));
    rows.forEach((tr) => {
      const tds = Array.from(tr.querySelectorAll("td"));
      const td = tds[schema[keySchemaIndex]?.idx ?? 0];
      const key = _norm(__poExtractCellValue(td));
      if (key) map.set(key, tr);
    });
    return map;
  }

  function __poFindRowKeyLoose(rowMap, key) {
    if (!key) return null;
    if (rowMap.has(key)) return key;

    const hits = [];
    for (const k of rowMap.keys()) {
      if (!k) continue;
      if (k.includes(key) || key.includes(k)) hits.push(k);
      if (hits.length > 2) break;
    }
    return hits.length === 1 ? hits[0] : null;
  }

  function __poSetNativeValue(input, value) {
    const v = value ?? "";
    const proto = input instanceof HTMLTextAreaElement ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) desc.set.call(input, v);
    else input.value = v;
  }

// ✅ Normaliza números vindos do Excel / CSV para o formato que os campos numéricos do PO-UI aceitam
// - "15,23"     -> "15.23"
// - "1.234,56"  -> "1234.56"
// - "1,234.56"  -> "1234.56"
// - remove espaços
// ===========================================================
// ✅ NÚMEROS (PT-BR): mantém vírgula como separador decimal
//    Ex.: "15,23" -> "15,23" | "15.23" -> "15,23" | "1.234,56" -> "1234,56"
// ===========================================================
const __poWait = (ms) => new Promise(r => setTimeout(r, ms));

function __poNormalizeNumericStringBR(raw) {
  let s = (raw ?? "").toString().trim();
  if (!s) return "";

  // remove espaços e tudo que não for dígito/separador/sinal
  s = s.replace(/\s+/g, "");
  let sign = "";
  if (s[0] === "-" || s[0] === "+") {
    sign = s[0];
    s = s.slice(1);
  }

  // mantém somente dígitos e separadores (.,)
  s = s.replace(/[^\d.,]/g, "");
  if (!s) return sign ? sign + "0" : "0";

  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");

  let decSep = null;

  if (lastComma >= 0 && lastDot >= 0) {
    // se ambos existem, o último normalmente é o decimal
    decSep = lastComma > lastDot ? "," : ".";
  } else if (lastComma >= 0) {
    // só vírgula: para notas, assume decimal
    decSep = ",";
  } else if (lastDot >= 0) {
    // só ponto: pode ser decimal OU milhar.
    // se houver vários pontos e o último tiver 3 dígitos, assume milhar.
    const parts = s.split(".");
    const last = parts[parts.length - 1] || "";
    const dotCount = parts.length - 1;
    if (dotCount >= 1 && last.length === 3 && parts.slice(0, -1).every(p => p.length > 0 && p.length <= 3)) {
      decSep = null; // milhar
    } else {
      decSep = ".";
    }
  }

  let intPart = s;
  let fracPart = "";

  if (decSep) {
    const idx = decSep === "," ? lastComma : lastDot;
    intPart = s.slice(0, idx);
    fracPart = s.slice(idx + 1);
  }

  // remove separadores de milhar do inteiro
  intPart = intPart.replace(/[.,]/g, "");
  // fração só dígitos
  fracPart = fracPart.replace(/[^\d]/g, "");

  // evita vazio
  if (!intPart) intPart = "0";

  let out = sign + intPart;
  if (fracPart !== "") out += "," + fracPart;

  return out;
}

function __poParseLooseNumber(raw) {
  const s0 = (raw ?? "").toString().trim();
  if (!s0) return NaN;

  let s = s0.replace(/\s+/g, "").replace(/[^\d.,\-+]/g, "");
  if (!s) return NaN;

  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");

  let decSep = null;
  if (lastComma >= 0 && lastDot >= 0) decSep = lastComma > lastDot ? "," : ".";
  else if (lastComma >= 0) decSep = ",";
  else if (lastDot >= 0) decSep = ".";

  if (decSep === ",") {
    s = s.replace(/\./g, "");
    s = s.replace(",", ".");
  } else if (decSep === ".") {
    s = s.replace(/,/g, "");
  } else {
    s = s.replace(/[.,]/g, "");
  }

  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}

function __poNumericEquivalent(a, b) {
  const na = __poParseLooseNumber(a);
  const nb = __poParseLooseNumber(b);
  if (!Number.isFinite(na) || !Number.isFinite(nb)) return false;
  return Math.abs(na - nb) < 1e-6;
}


function __poLooksNumericLike(raw) {
  const s = (raw ?? "").toString().trim();
  if (!s) return false;
  return /^[+-]?[0-9][0-9\s.,]*$/.test(s) && /\d/.test(s);
}

  async function __poSetCellValue(td, value) {
    if (!td) return false;

    try {
      td.scrollIntoView({ block: "nearest", inline: "center" });
    } catch {}

    const want = (value ?? "").toString().trim();

let input =
  td.querySelector("input, textarea") ||
  td.querySelector("po-input input, po-number input, po-decimal input, po-field-container input") ||
  td.querySelector(".po-field-container input");

if (!input) {
  // tenta ativar edição
  td.click();
  td.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
  await wait(60);

  input =
    td.querySelector("input, textarea") ||
    td.querySelector("po-input input, po-number input, po-decimal input, po-field-container input") ||
    document.querySelector(".po-popup-content input, .po-modal input, .cdk-overlay-container input");
}

if (!input) return false;


// ✅ Ajuste de decimais (PT-BR) + escrita robusta em inputs mascarados do PO-UI
const inputMode = (input.getAttribute("inputmode") || "").toLowerCase();
const numericLike = __poLooksNumericLike(want) && (
  input.type === "number" ||
  inputMode === "numeric" ||
  inputMode === "decimal" ||
  inputMode === "tel" ||
  input.getAttribute("data-po-number") === "true" ||
  /[.,]/.test(want)
);

let v = want;
if (numericLike) v = __poNormalizeNumericStringBR(want);

// 1) tentativa padrão (set native value + eventos)
input.focus();
__poSetNativeValue(input, "");
input.dispatchEvent(new Event("input", { bubbles: true }));
await __poWait(0);

__poSetNativeValue(input, v);
input.dispatchEvent(new Event("input", { bubbles: true }));
input.dispatchEvent(new Event("change", { bubbles: true }));
input.dispatchEvent(new Event("blur", { bubbles: true }));

// 2) fallback: se era número e o componente “comeu” a vírgula/ponto (ex.: 15,23 virando 1523,00),
// digita caractere por caractere para o mascaramento capturar corretamente.
if (numericLike) {
  const got = (input.value ?? "").toString();
  if (!__poNumericEquivalent(got, v)) {
    __poSetNativeValue(input, "");
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await __poWait(0);

    for (const ch of v) {
      __poSetNativeValue(input, (input.value ?? "") + ch);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      await __poWait(8);
    }

    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(new Event("blur", { bubbles: true }));
  }
}
// alguns componentes precisam do Enter para "confirmar"
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", code: "Enter", bubbles: true }));

    return true;
  }

  async function __poApplyNotasToTable(imported, setStatusFn) {
    const table = __poGetNotasTable();
    if (!table) return { ok: false, msg: "Tabela de notas não encontrada nesta tela." };

    // ✅ garante que todas as linhas estejam carregadas antes de mapear
    if (setStatusFn) setStatusFn("Carregando todas as linhas da tabela...");
    await __poExpandAllRowsForNotasTable(table, null, null);

    const schema = __poGetTableSchema(table);
    if (!schema) return { ok: false, msg: "Não consegui ler a estrutura (thead) da tabela de notas." };

    const headersNorm = (imported.headers || []).map((h) => _norm(h));
    const keyIdxImported = __poFindKeyColumnIndex(headersNorm);
    const keyIdxSchema = __poFindKeyColumnIndexInSchema(schema);

    const schemaNames = schema.map((c) => c.label || c.key || "");
    const colMap = new Map(); // importedIdx => schemaIdx

    for (let i = 0; i < headersNorm.length; i++) {
      if (i === keyIdxImported) continue;
      const h = headersNorm[i];
      if (!h) continue;

      let si = schemaNames.findIndex((n) => n === h || (n && h && (n.includes(h) || h.includes(n))));
      if (si < 0) {
        si = schemaNames.findIndex((n) => (n || "").replace(/\s/g, "") === h.replace(/\s/g, ""));
      }
      if (si >= 0) colMap.set(i, si);
    }

    if (colMap.size === 0) {
      return {
        ok: false,
        msg:
          "Não consegui mapear colunas do arquivo para a tabela.\n" +
          "Dica: mantenha os títulos da primeira linha iguais aos títulos da tela (ex.: N1, N2, Média, etc.).",
      };
    }

    const rowMap = __poBuildRowMapByKey(table, schema, keyIdxSchema);

    let totalValidKeys = 0;
    let matchedRows = 0;
    let writtenCells = 0;
    let failedCells = 0;

    for (let r = 0; r < imported.data.length; r++) {
      const row = imported.data[r] || [];
      const keyRaw = (row[keyIdxImported] ?? "").toString();
      const key = _norm(keyRaw);
      if (!key) continue;

      totalValidKeys++;

      let tr = rowMap.get(key);
      if (!tr) {
        const looseKey = __poFindRowKeyLoose(rowMap, key);
        if (looseKey) tr = rowMap.get(looseKey);
      }
      if (!tr) continue;

      matchedRows++;
      const tds = Array.from(tr.querySelectorAll("td"));

      for (const [impColIdx, schemaIdx] of colMap.entries()) {
        const col = schema[schemaIdx];
        const td = tds[col.idx];
        if (!td) {
          failedCells++;
          continue;
        }

        const val = (row[impColIdx] ?? "").toString();
        const ok = await __poSetCellValue(td, val);
        if (ok) writtenCells++;
        else failedCells++;

        // pequeno respiro p/ UI
        if ((writtenCells + failedCells) % 12 === 0) await wait(10);
      }

      if (setStatusFn && r % 8 === 0) {
        setStatusFn(
          `Importando...\nLinhas: ${r + 1}/${imported.data.length}\nEncontradas na tela: ${matchedRows}\nCélulas escritas: ${writtenCells}\nFalhas: ${failedCells}`
        );
        await wait(10);
      }
    }

    return {
      ok: true,
      msg:
        `Import concluído ✅\n` +
        `Linhas no arquivo: ${imported.data.length}\n` +
        `Linhas com chave válida: ${totalValidKeys}\n` +
        `Linhas encontradas na tela: ${matchedRows}\n` +
        `Células escritas: ${writtenCells}\n` +
        `Falhas: ${failedCells}\n\n` +
        `Obs: colunas não editáveis na tela podem falhar.`,
    };
  }

  function __poRemoveNotasPanel() {
    const p = document.getElementById(PO_NOTAS_PANEL_ID);
    if (p) p.remove();
  }

  function __poEnsureNotasPanel() {
    if (document.getElementById(PO_NOTAS_PANEL_ID)) return;

    __poInjectNotasStylesOnce();

    const panel = document.createElement("div");
    panel.id = PO_NOTAS_PANEL_ID;
    panel.innerHTML = `
      <div class="po-notas-head">
        <div class="po-notas-title">Notas (Excel)</div>
        <div class="po-notas-close" title="Fechar (o botão de reabrir fica no painel principal)">×</div>
      </div>

      <div class="po-notas-body">
        <div class="po-notas-row">
          <button class="po-notas-btn orange" id="poNotasDownloadBtn">⬇ Baixar Excel (.xlsx)</button>
          <button class="po-notas-btn green" id="poNotasImportBtn">⬆ Importar Excel</button>
        </div>

        <div class="po-notas-drop" id="poNotasDrop">
          Arraste o Excel aqui e solte<br/>
          <span style="opacity:.8">(.xlsx / .csv)</span>
        </div>

        <div class="po-notas-status" id="poNotasStatus">Pronto.</div>
        <input type="file" id="poNotasFileInput" accept=".xlsx,.xlsm,.xltx,.csv,.txt,.xls" style="display:none;" />
      </div>
    `;

    document.body.appendChild(panel);

    const statusEl = panel.querySelector("#poNotasStatus");
    const setStatus = (t) => {
      if (statusEl) statusEl.textContent = t;
    };

    // close/hide => mostra launcher no painel principal
    panel.querySelector(".po-notas-close").addEventListener("click", () => {
      __poNotasPanelHiddenSave(true);
      __poRemoveNotasPanel();
      __poNotasPanelAutoToggle(); // atualiza launcher
    });

    // download
    panel.querySelector("#poNotasDownloadBtn").addEventListener("click", async () => {
      setStatus("Verificando tabela...");
      const table = __poGetNotasTable();
      if (!table) {
        setStatus("Tabela de notas não encontrada nesta tela.");
        return;
      }
      setStatus("Exportando...");
      await __poExportNotasExcel(setStatus);
      setStatus("Pronto ✅");
    });

    // import (file input)
    const input = panel.querySelector("#poNotasFileInput");
    panel.querySelector("#poNotasImportBtn").addEventListener("click", () => input.click());

    input.addEventListener("change", async () => {
      const file = input.files?.[0];
      input.value = "";
      if (!file) return;

      setStatus("Lendo arquivo...");
      let parsed = null;
      try {
        parsed = await __poReadNotasFile(file);
      } catch (e) {
        setStatus("Falha ao ler arquivo: " + (e?.message || e));
        return;
      }

      if (!parsed) {
        setStatus("Não consegui ler este arquivo.");
        return;
      }

      setStatus("Aplicando valores na tabela (carregando linhas)...");
      const res = await __poApplyNotasToTable(parsed, setStatus);
      setStatus(res.msg || (res.ok ? "Import finalizado ✅" : "Falha ao importar."));
    });

    // drag & drop
    const drop = panel.querySelector("#poNotasDrop");
    const stop = (e) => {
      e.preventDefault();
      e.stopPropagation();
    };

    ["dragenter", "dragover"].forEach((ev) =>
      drop.addEventListener(ev, (e) => {
        stop(e);
        drop.classList.add("dragover");
      })
    );
    ["dragleave", "drop"].forEach((ev) =>
      drop.addEventListener(ev, (e) => {
        stop(e);
        drop.classList.remove("dragover");
      })
    );

    drop.addEventListener("drop", async (e) => {
      const file = e.dataTransfer?.files?.[0];
      if (!file) return;

      setStatus("Lendo arquivo...");
      let parsed = null;
      try {
        parsed = await __poReadNotasFile(file);
      } catch (err) {
        setStatus("Falha ao ler arquivo: " + (err?.message || err));
        return;
      }

      if (!parsed) {
        setStatus("Não consegui ler este arquivo.");
        return;
      }

      setStatus("Aplicando valores na tabela (carregando linhas)...");
      const res = await __poApplyNotasToTable(parsed, setStatus);
      setStatus(res.msg || (res.ok ? "Import finalizado ✅" : "Falha ao importar."));
    });
  }

  // ✅ botão no painel principal para reabrir Notas
  function __poUpdateNotasLauncher(show) {
    const row = document.getElementById("poNotasLauncherRow");
    if (!row) return;
    row.style.display = show ? "flex" : "none";
  }

  function __poNotasPanelAutoToggle(forceShow = false) {
    const table = __poGetNotasTable();

    if (forceShow) __poNotasPanelHiddenSave(false);

    if (table) {
      if (__poNotasPanelHiddenLoad()) {
        __poUpdateNotasLauncher(true);
        __poRemoveNotasPanel();
      } else {
        __poUpdateNotasLauncher(false);
        __poEnsureNotasPanel();
      }
    } else {
      __poUpdateNotasLauncher(false);
      __poRemoveNotasPanel();
    }
  }

  // ============================================
  // ✅ ARQUIVAMENTO (LIXEIRA - mantém igual)
  // ============================================
  function generateRowHash(row) {
    return row.innerText.trim().replace(/\s+/g, "").substring(0, 150);
  }

  function getColumnIndexByName(table, name) {
    if (!table) return -1;
    const th = table.querySelector(`thead th[data-po-table-column-name="${CSS.escape(name)}"]`);
    if (!th || !th.parentNode) return -1;
    return Array.from(th.parentNode.children).indexOf(th);
  }

  function extractRichData(table, row) {
    const tds = row.querySelectorAll("td");
    if (tds.length === 0) return { disciplina: "...", turma: "", curso: "" };

    const idxCurso = getColumnIndexByName(table, "curso");
    const idxTurma = getColumnIndexByName(table, "cód. turma");
    const idxDisc = getColumnIndexByName(table, "disciplina");

    let curso = idxCurso > -1 && tds[idxCurso] ? tds[idxCurso].innerText : "";
    let turma = idxTurma > -1 && tds[idxTurma] ? tds[idxTurma].innerText : "";
    let disc = idxDisc > -1 && tds[idxDisc] ? tds[idxDisc].innerText : "";

    if (!disc && tds.length > 5) disc = tds[5].innerText;
    if (!turma && tds.length > 4) turma = tds[4].innerText;
    if (!curso && tds.length > 2) curso = tds[2].innerText;

    return {
      disciplina: (disc || "").trim() || "Item sem nome",
      turma: (turma || "").trim() || "",
      curso: (curso || "").trim() || "",
    };
  }

  function runArchivedCleanup() {
    const table = getArchiveTable();

    document.querySelectorAll(".po-archive-check-container").forEach((el) => {
      if (!table || !el.closest("table") || el.closest("table") !== table) el.remove();
    });

    if (!table) return;

    const archivedData = loadArchivedData();
    const archivedIds = new Set(archivedData.map((item) => String(item.id)));

    const rows = table.querySelectorAll("tbody tr");

    rows.forEach((row) => {
      const firstTd = row.querySelector("td");
      if (!firstTd) return;

      const id = generateRowHash(row);
      if (!id) return;

      if (archivedIds.has(String(id))) {
        row.remove();
        return;
      }

      if (!row.querySelector(".po-archive-check-container")) {
        if (getComputedStyle(firstTd).position === "static") {
          firstTd.style.position = "relative";
        }
        const checkContainer = document.createElement("div");
        checkContainer.className = "po-archive-check-container";
        checkContainer.innerHTML = `<input type="checkbox" class="po-archive-check" title="Marcar para Arquivar" />`;
        firstTd.insertBefore(checkContainer, firstTd.firstChild);
      }
    });

    renderArchivedListInPanel();
  }

  function archiveSelectedRows() {
    const table = getArchiveTable();
    if (!table) {
      alert("Tabela do Diário de Classe não encontrada nesta tela.");
      return;
    }

    const rows = table.querySelectorAll("tbody tr");
    const currentList = loadArchivedData();
    const currentIds = new Set(currentList.map((i) => String(i.id)));
    let count = 0;

    rows.forEach((row) => {
      const checkbox = row.querySelector(".po-archive-check");
      if (checkbox && checkbox.checked) {
        const id = generateRowHash(row);
        const richData = extractRichData(table, row);

        if (id && !currentIds.has(String(id))) {
          currentList.push({
            id: id,
            disciplina: richData.disciplina,
            turma: richData.turma,
            curso: richData.curso,
          });
          currentIds.add(String(id));
          row.remove();
          count++;
        }
      }
    });

    if (count > 0) {
      saveArchivedData(currentList);
      renderArchivedListInPanel();
    } else {
      alert("Selecione itens usando a caixinha na esquerda da tabela.");
    }
  }

  function restoreSelectedArchived() {
    const checks = document.querySelectorAll(".po-restore-check:checked");
    if (checks.length === 0) {
      alert("Selecione itens na lista acima para restaurar.");
      return;
    }

    if (!confirm(`Restaurar ${checks.length} itens?`)) return;

    const idsToRestore = new Set(Array.from(checks).map((c) => c.value));
    const currentList = loadArchivedData();
    const newList = currentList.filter((item) => !idsToRestore.has(String(item.id)));

    saveArchivedData(newList);
    localStorage.setItem(STORAGE_KEY_EXPAND, "true");
    location.reload();
  }

  function restoreAllArchived() {
    if (!confirm("Restaurar TUDO?")) return;
    saveArchivedData([]);
    localStorage.setItem(STORAGE_KEY_EXPAND, "true");
    location.reload();
  }

  // =======================
  // PAINEL (UI) - inferior esquerdo
  // =======================
  let _lastRenderedArchivedHash = "";

  function renderArchivedListInPanel() {
    const listContainer = document.getElementById("poArchivedItemsList");
    const countLabel = document.getElementById("poArchivedCountLabel");
    const btnRestoreAll = document.getElementById("poRestoreAllBtn");
    const btnRestoreSel = document.getElementById("poRestoreSelectedBtn");

    if (!listContainer) return;

    const data = loadArchivedData();

    const term = _currentSearchTerm.toLowerCase();
    const filteredData = data.filter((item) => {
      if (!term) return true;
      const text = ((item.disciplina || "") + (item.turma || "") + (item.curso || "") + (item.name || "")).toLowerCase();
      return text.includes(term);
    });

    const currentHash = JSON.stringify(data) + "||" + term;
    if (currentHash === _lastRenderedArchivedHash && listContainer.children.length > 0) return;
    _lastRenderedArchivedHash = currentHash;

    if (countLabel) {
      if (term) countLabel.textContent = `(${filteredData.length} de ${data.length})`;
      else countLabel.textContent = `(${data.length})`;
    }

    const hasData = filteredData.length > 0;
    if (btnRestoreAll) {
      btnRestoreAll.disabled = data.length === 0;
      btnRestoreAll.style.opacity = data.length === 0 ? "0.5" : "1";
    }
    if (btnRestoreSel) {
      btnRestoreSel.disabled = !hasData;
      btnRestoreSel.style.opacity = hasData ? "1" : "0.5";
    }

    listContainer.innerHTML = "";

    if (data.length === 0) {
      listContainer.innerHTML = `<div style="padding:8px; opacity:0.6; font-style:italic; font-size:11px; text-align:center;">Lixeira vazia.</div>`;
      return;
    }

    if (data.length > 0 && filteredData.length === 0) {
      listContainer.innerHTML = `<div style="padding:8px; opacity:0.6; font-style:italic; font-size:11px; text-align:center;">Nenhum item encontrado para "${term}".</div>`;
      return;
    }

    filteredData.forEach((item) => {
      const div = document.createElement("div");
      div.className = "po-archived-item-row";

      const disc = item.disciplina || item.name || "Item Antigo";
      const turma = item.turma || "";
      const curso = item.curso || "";

      div.innerHTML = `
        <label style="display:flex; align-items:flex-start; gap:8px; width:100%; cursor:pointer; user-select:none;">
          <input type="checkbox" class="po-restore-check" value="${item.id}" style="margin-top:4px;" />
          <div style="flex:1; overflow:hidden;">
            <div class="po-archived-discipline" title="${disc}">${disc}</div>
            ${turma ? `<div class="po-archived-code">${turma}</div>` : ""}
            ${curso ? `<div class="po-archived-course" title="${curso}">${curso}</div>` : ""}
          </div>
        </label>
      `;
      listContainer.appendChild(div);
    });
  }

  function renderHiddenList() {
    const box = document.getElementById("poHiddenList");
    if (!box) return;

    const hidden = loadHidden();
    box.innerHTML = "";

    if (hidden.size === 0) {
      const div = document.createElement("div");
      div.className = "small";
      div.textContent = "Nenhuma coluna oculta.";
      box.appendChild(div);
      return;
    }

    for (const colName of hidden) {
      const th = document.querySelector(`${TH_SELECTOR}[data-po-table-column-name="${CSS.escape(colName)}"]`);
      const label = th ? (th.innerText || colName).trim() : colName;

      const item = document.createElement("div");
      item.className = "po-hidden-item";

      const pill = document.createElement("div");
      pill.className = "po-pill";
      pill.title = label;
      pill.textContent = label;

      const btn = document.createElement("button");
      btn.className = "btn";
      btn.textContent = "Mostrar";
      btn.addEventListener("click", () => {
        unhideColumn(colName);
        renderHiddenList();
      });

      item.appendChild(pill);
      item.appendChild(btn);
      box.appendChild(item);
    }
  }

  function ensurePanel() {
    if (document.getElementById("poFontPanel")) return;

    const panel = document.createElement("div");
    panel.id = "poFontPanel";

    panel.style.left = `${PANEL_LEFT}px`;
    panel.style.bottom = `${PANEL_BOTTOM}px`;

    panel.innerHTML = `
      <div id="poPanelMinBtn" title="Minimizar / Restaurar">–</div>

      <div class="panel-body">
        <div class="row" style="margin-top:5px;">
          <div class="title">Fonte</div>
          <div class="small" id="poFontValue">${loadFontPx()}px</div>
        </div>

        <div class="row">
          <input id="poFontRange" type="range" min="${FONT_MIN}" max="${FONT_MAX}" step="1" />
        </div>

        <div class="small" style="margin-top:8px; border-top:1px solid rgba(255,255,255,0.2); padding-top:5px;">Layout</div>
        <div class="row">
          <div style="display:flex; align-items:center; gap:5px;">
            <div class="small" id="poLockState">—</div>
            <button class="btn" id="poLockToggle" title="Travado/Original">🔓</button>
          </div>
          
        </div>

        <div id="poHiddenList"></div>

        <!-- ✅ Launcher Notas (aparece somente na tela de notas quando o painel foi fechado) -->
        <div class="row" id="poNotasLauncherRow" style="display:none; margin-top:8px;">
          <button class="btn" id="poNotasLauncherBtn" title="Reabrir painel Notas (Excel)">📊</button>
          <div class="small" style="opacity:.85;">Reabrir Notas</div>
        </div>

        <!-- LIXEIRA (mantida) -->
        <div class="small" style="margin-top:12px; border-top:1px solid rgba(255,255,255,0.2); padding-top:5px; font-weight:bold; display:flex; justify-content:space-between; align-items:center;">
          <span>UNIDADES CURRICULARES ARQUIVADAS <span id="poArchivedCountLabel">(0)</span></span>
          <input type="text" id="poArchivedSearch" placeholder="🔍 Buscar..." style="width: 140px; padding: 2px 5px; border-radius: 4px; border: none; font-size: 11px; color:#333;" />
        </div>

        <div class="row">
          <button class="btn po-btn-red" id="poArchiveSelectedBtn" style="width:100%; font-size:15px; margin-top:5px;">Arquivar Marcados 📦</button>
        </div>

        <div id="poArchivedItemsList" class="po-archived-list-container"></div>

        <div class="row" style="margin-top:8px; gap:5px;">
          <button class="btn po-btn-green" id="poRestoreSelectedBtn" style="flex:1; font-size:15px;">Restaurar Marcados</button>
          <button class="btn po-btn-orange" id="poRestoreAllBtn" style="flex:1; font-size:15px;">Restaurar Tudo</button>
        </div>
      </div>
    `;

    document.body.appendChild(panel);

    const savedSize = loadPanelSize();
    applyPanelSize(panel, savedSize);
    installPanelResizers(panel);

    const range = panel.querySelector("#poFontRange");
    range.addEventListener("input", () => applyFont(Number(range.value)));

    panel.querySelector("#poArchivedSearch").addEventListener("input", (e) => {
      _currentSearchTerm = e.target.value;
      renderArchivedListInPanel();
    });

    // ✅ launcher notas
    panel.querySelector("#poNotasLauncherBtn").addEventListener("click", () => {
      __poNotasPanelHiddenSave(false);
      __poNotasPanelAutoToggle(true);
    });

    const minBtn = panel.querySelector("#poPanelMinBtn");

    function setMinimizedState(min) {
      const isMin = !!min;

      if (isMin) {
        if (!panel.classList.contains("minimized")) {
          const r = panel.getBoundingClientRect();
          const clamped = clampPanelSizeToViewport(panel, r.width, r.height);
          savePanelSize(clamped.w, clamped.h);
        }

        panel.classList.add("minimized");
        saveMinimized(true);

        panel.style.setProperty("width", "40px", "important");
        panel.style.setProperty("height", "40px", "important");
        panel.style.setProperty("max-width", "40px", "important");
        panel.style.setProperty("max-height", "40px", "important");

        minBtn.textContent = "+";
        minBtn.title = "Restaurar";
        return;
      }

      panel.classList.remove("minimized");
      saveMinimized(false);

      panel.style.removeProperty("width");
      panel.style.removeProperty("height");
      panel.style.removeProperty("max-width");
      panel.style.removeProperty("max-height");

      minBtn.textContent = "–";
      minBtn.title = "Minimizar";

      applyPanelSize(panel, loadPanelSize());
    }

    minBtn.addEventListener("click", () => setMinimizedState(!panel.classList.contains("minimized")));
    //panel.querySelector("#poRefreshHidden").addEventListener("click", renderHiddenList);
    panel.querySelector("#poLockToggle").addEventListener("click", () => {
      const next = !loadLock();
      if (next) saveBoxWidth(null);
      applyLockState(next);
    });

    panel.querySelector("#poArchiveSelectedBtn").addEventListener("click", archiveSelectedRows);
    panel.querySelector("#poRestoreSelectedBtn").addEventListener("click", restoreSelectedArchived);
    panel.querySelector("#poRestoreAllBtn").addEventListener("click", restoreAllArchived);

    setMinimizedState(loadMinimized());
    renderHiddenList();
    updateLockUI();
    applyFont(loadFontPx());
    renderArchivedListInPanel();

    // atualiza o launcher assim que o painel principal existir
    __poNotasPanelAutoToggle();
  }

  // =======================
  // COLUNAS (resizer/hide)
  // =======================
  function setThWidth(th, px) {
    const v = `${px}px`;
    th.style.setProperty("width", v, "important");
    th.style.setProperty("min-width", v, "important");
    th.style.setProperty("max-width", v, "important");
  }
  function measureTextWidth(text, sampleEl) {
    const meas = document.createElement("span");
    const cs = getComputedStyle(sampleEl);
    meas.textContent = text || "";
    meas.style.cssText = "position:fixed;left:-9999px;visibility:hidden;white-space:nowrap;";
    meas.style.font = cs.font;
    meas.style.letterSpacing = cs.letterSpacing;
    meas.style.textTransform = cs.textTransform;
    document.body.appendChild(meas);
    const w = meas.getBoundingClientRect().width;
    meas.remove();
    return w;
  }
  function getThIndex(th) {
    const row = th.parentElement;
    if (!row) return -1;
    return Array.from(row.querySelectorAll("th")).indexOf(th);
  }
  function autoFitColumn(th) {
    const table = th.closest("table");
    if (!table) return;
    const idx = getThIndex(th);
    if (idx < 0) return;

    const colName = th.getAttribute("data-po-table-column-name") || `col_${idx}`;
    let maxW = measureTextWidth((th.innerText || "").trim(), th);

    Array.from(table.querySelectorAll("tbody tr")).forEach((r) => {
      const td = r.querySelectorAll("td")[idx];
      if (td) {
        const w = measureTextWidth((td.innerText || "").trim(), td);
        if (w > maxW) maxW = w;
      }
    });

    const target = clamp(Math.round(maxW + EXTRA_PAD), MIN_W, MAX_W);
    setThWidth(th, target);
    const map = loadWidths();
    map[colName] = target;
    saveWidths(map);
    syncBoxWidthToColumns();
  }
  function hideColumn(colName) {
    const h = loadHidden();
    h.add(colName);
    saveHidden(h);
    clearHiddenClasses();
    applyHiddenState();
    syncBoxWidthToColumns();
  }
  function unhideColumn(colName) {
    const h = loadHidden();
    h.delete(colName);
    saveHidden(h);
    clearHiddenClasses();
    applyHiddenState();
    syncBoxWidthToColumns();
  }
  function clearHiddenClasses() {
    document.querySelectorAll(".po-col-hidden").forEach((el) => el.classList.remove("po-col-hidden"));
  }
  function applyHiddenState() {
    const table = document.querySelector(TABLE_SELECTOR);
    if (!table) return;

    const hidden = loadHidden();
    const ths = Array.from(table.querySelectorAll("thead th"));

    hidden.forEach((colName) => {
      const th = table.querySelector(`${TH_SELECTOR}[data-po-table-column-name="${CSS.escape(colName)}"]`);
      if (!th) return;
      const idx = ths.indexOf(th);
      if (idx < 0) return;

      th.classList.add("po-col-hidden");
      Array.from(table.querySelectorAll("tbody tr")).forEach((r) => {
        const td = r.querySelectorAll("td")[idx];
        if (td) td.classList.add("po-col-hidden");
      });
    });
  }
  function getGuide() {
    let guide = document.getElementById("poResizeGuide");
    if (!guide) {
      guide = document.createElement("div");
      guide.id = "poResizeGuide";
      document.body.appendChild(guide);
    }
    return guide;
  }
  function installHeaderControls() {
    const table = document.querySelector(TABLE_SELECTOR);
    if (!table) return false;

    const ths = document.querySelectorAll(TH_SELECTOR);
    const saved = loadWidths();
    const hidden = loadHidden();

    ths.forEach((th) => {
      const colName = th.getAttribute("data-po-table-column-name");
      if (!colName) return;

      if (saved[colName]) setThWidth(th, saved[colName]);

      if (th.dataset.poHideInstalled !== "1") {
        const btn = document.createElement("div");
        btn.className = "po-col-hide-btn";
        btn.textContent = hidden.has(colName) ? "🚫" : "👁";
        btn.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          if (loadHidden().has(colName)) unhideColumn(colName);
          else hideColumn(colName);
          renderHiddenList();
        });
        th.appendChild(btn);
        th.dataset.poHideInstalled = "1";
      }

      if (th.dataset.poResizerInstalled !== "1") {
        const handle = document.createElement("div");
        handle.className = "po-col-resizer";
        th.appendChild(handle);

        handle.addEventListener("dblclick", (e) => {
          e.preventDefault();
          e.stopPropagation();
          autoFitColumn(th);
        });

        handle.addEventListener("pointerdown", (e) => {
          e.preventDefault();
          e.stopPropagation();
          handle.setPointerCapture(e.pointerId);

          const guide = getGuide();
          const startX = e.clientX;
          const startW = th.getBoundingClientRect().width;
          const thLeft = th.getBoundingClientRect().left;

          guide.style.display = "block";
          guide.style.transform = `translateX(${e.clientX}px)`;
          document.body.classList.add("po-resizing");

          let finalW = startW;

          const onMove = (ev) => {
            const dx = ev.clientX - startX;
            finalW = clamp(Math.round(startW + dx), MIN_W, MAX_W);
            guide.style.transform = `translateX(${thLeft + finalW}px)`;
          };

          const onUp = () => {
            document.removeEventListener("pointermove", onMove, true);
            document.removeEventListener("pointerup", onUp, true);
            document.body.classList.remove("po-resizing");
            guide.style.display = "none";

            setThWidth(th, finalW);
            const map = loadWidths();
            map[colName] = finalW;
            saveWidths(map);
            syncBoxWidthToColumns();
          };

          document.addEventListener("pointermove", onMove, true);
          document.addEventListener("pointerup", onUp, true);
        });

        th.dataset.poResizerInstalled = "1";
      }
    });

    clearHiddenClasses();
    applyHiddenState();

    if (loadLock()) {
      const sb = loadBoxWidth();
      if (sb && sb > 0) setWrapperFixedWidth(sb);
      else syncBoxWidthToColumns();
    }

    return true;
  }

  // =======================
  // AULAS INDICATORS (mantém igual)
  // =======================
  function getPanelTextareas(panel) {
    const arr = Array.from(panel.querySelectorAll("textarea.po-textarea"));
    return { previsto: arr[0] || null, realizado: arr[1] || null };
  }
  function updateAulasIndicators(panel) {
    if (!panel) return;
    const { previsto, realizado } = getPanelTextareas(panel);
    const pF = !!(previsto && (previsto.value || "").trim().length > 0);
    const rF = !!(realizado && (realizado.value || "").trim().length > 0);

    const setChip = (k, f) => {
      const c = panel.querySelector(`.po-ext-aulas-chip[data-kind="${k}"]`);
      if (c) {
        const cb = c.querySelector("input");
        if (cb) cb.checked = f;
        c.classList.toggle("filled", f);
      }
    };

    setChip("previsto", pF);
    setChip("realizado", rF);
  }
  function ensureAulasIndicatorsForPanel(panel) {
    if (!panel || panel.dataset.poAulasIndicator2 === "1") return;
    const header = panel.querySelector("mat-expansion-panel-header");
    if (!header) return;

    const t =
      header.querySelector("mat-panel-title") ||
      header.querySelector(".mat-expansion-panel-header-title") ||
      header.querySelector("span.mat-content");

    const w = document.createElement("span");
    w.className = "po-ext-aulas-indicator";
    w.innerHTML = `
      <span class="po-ext-aulas-chip" data-kind="previsto"><input type="checkbox" disabled /><span class="lbl">Previsto</span></span>
      <span class="po-ext-aulas-chip" data-kind="realizado"><input type="checkbox" disabled /><span class="lbl">Realizado</span></span>
    `;

    if (t && t.parentElement) t.parentElement.insertBefore(w, t);
    else header.insertBefore(w, header.firstChild);

    panel.dataset.poAulasIndicator2 = "1";
    updateAulasIndicators(panel);
  }
  function installAulasIndicators() {
    document.querySelectorAll("mat-expansion-panel.mat-expansion-panel").forEach(ensureAulasIndicatorsForPanel);
  }
  function installAulasDelegatedListenersOnce() {
    if (window.__poAulasDelegated2) return;
    window.__poAulasDelegated2 = true;

    const h = (e) => {
      const t = e.target;
      if (t instanceof HTMLTextAreaElement && t.classList.contains("po-textarea")) {
        updateAulasIndicators(t.closest("mat-expansion-panel.mat-expansion-panel"));
      }
    };

    document.addEventListener("input", h, true);
    document.addEventListener("change", h, true);
  }

  // =======================
  // FIX CHECKBOX (mantém igual)
  // =======================
  function isolateCheckboxesInHeaders() {
    const targets = document.querySelectorAll('mat-expansion-panel-header po-checkbox:not([data-po-fix="7"])');
    targets.forEach((cb) => {
      cb.dataset.poFix = "7";
      cb.addEventListener("click", (e) => {
        const isLabel =
          e.target.closest("po-label") || e.target.classList.contains("po-label") || e.target.tagName === "LABEL";
        if (isLabel) {
          e.preventDefault();
          return;
        }
        e.stopPropagation();
      });
      const outlines = cb.querySelectorAll(".po-checkbox-outline, .po-checkbox-input");
      outlines.forEach((o) => o.addEventListener("mousedown", (e) => e.stopPropagation()));
      cb.addEventListener("keydown", (e) => {
        if (e.key === " " || e.code === "Space") e.stopPropagation();
      });
    });
  }

  function highlightByDateChange() {
    const panels = document.querySelectorAll("mat-expansion-panel.mat-expansion-panel");
    panels.forEach((p) => p.classList.remove("po-highlight-first"));
    let lastSeenDate = null;

    panels.forEach((panel) => {
      const header = panel.querySelector("mat-expansion-panel-header");
      if (!header) return;
      const text = (header.innerText || "").trim();
      const match = text.match(REGEX_DATE);
      if (match) {
        const currentDate = match[1];
        if (currentDate !== lastSeenDate) {
          panel.classList.add("po-highlight-first");
          lastSeenDate = currentDate;
        }
      }
    });
  }

  // =======================
  // MASTER TOOLS (mantém igual)
  // =======================
  function getPrevistoText(panel) {
    const textareas = panel.querySelectorAll("textarea.po-textarea, textarea");
    if (textareas.length > 0) return textareas[0].value || "";
    return "";
  }
  function injectTextIntoPanel(panel, text) {
    const textareas = panel.querySelectorAll("textarea.po-textarea, textarea");
    if (textareas.length > 0) {
      const target = textareas[0];
      if (target.value === text) return false;
      target.value = text;
      target.dispatchEvent(new Event("input", { bubbles: true }));
      target.dispatchEvent(new Event("change", { bubbles: true }));
      target.dispatchEvent(new Event("blur", { bubbles: true }));
      return true;
    }
    return false;
  }
  function clearContentFromPanel(panel) {
    const textareas = panel.querySelectorAll("textarea.po-textarea, textarea");
    let changed = false;
    textareas.forEach((t) => {
      if (t.value !== "") {
        t.value = "";
        t.dispatchEvent(new Event("input", { bubbles: true }));
        t.dispatchEvent(new Event("change", { bubbles: true }));
        t.dispatchEvent(new Event("blur", { bubbles: true }));
        changed = true;
      }
    });
    return changed;
  }
  function getDateFromPanel(panel) {
    const header = panel.querySelector("mat-expansion-panel-header");
    if (!header) return null;
    const text = (header.innerText || "").trim();
    const match = text.match(REGEX_DATE);
    return match ? match[1] : null;
  }

  function runCascadeCopy(statusBtn) {
    statusBtn.disabled = true;
    statusBtn.textContent = "Processando...";

    const panels = Array.from(document.querySelectorAll("mat-expansion-panel.mat-expansion-panel"));
    const selectedPanels = panels.filter((panel) => {
      const header = panel.querySelector("mat-expansion-panel-header");
      if (!header) return false;
      const cb = header.querySelector("po-checkbox");
      if (!cb) return false;
      const ariaEl = cb.querySelector('[role="checkbox"]');
      return ariaEl && ariaEl.getAttribute("aria-checked") === "true";
    });

    if (selectedPanels.length < 2) {
      alert("Selecione pelo menos 2 aulas para replicar em cascata.");
      statusBtn.disabled = false;
      statusBtn.textContent = "Replicar em Cascata (Dia)";
      return;
    }

    let activeText = null;
    let lastDate = null;
    let count = 0;

    selectedPanels.forEach((panel, index) => {
      const currentDate = getDateFromPanel(panel);
      const isSource = index === 0 || currentDate === null || currentDate !== lastDate;

      if (isSource) {
        activeText = getPrevistoText(panel);
        if (currentDate) lastDate = currentDate;
      } else {
        if (activeText !== null) {
          const changed = injectTextIntoPanel(panel, activeText);
          if (changed) {
            count++;
            updateAulasIndicators(panel);
          }
        }
      }
    });

    statusBtn.textContent = `${count} réplicas realizadas!`;
    setTimeout(() => {
      statusBtn.disabled = false;
      statusBtn.textContent = "Replicar em Cascata (Dia)";
    }, 2000);
  }

  function runClearContent(statusBtn) {
    statusBtn.disabled = true;
    statusBtn.textContent = "Limpando...";

    const panels = Array.from(document.querySelectorAll("mat-expansion-panel.mat-expansion-panel"));
    const selectedPanels = panels.filter((panel) => {
      const header = panel.querySelector("mat-expansion-panel-header");
      if (!header) return false;
      const cb = header.querySelector("po-checkbox");
      if (!cb) return false;
      const ariaEl = cb.querySelector('[role="checkbox"]');
      return ariaEl && ariaEl.getAttribute("aria-checked") === "true";
    });

    if (selectedPanels.length === 0) {
      statusBtn.textContent = "Selecione algo!";
      setTimeout(() => {
        statusBtn.disabled = false;
        statusBtn.textContent = "Limpar Conteúdos";
      }, 2000);
      return;
    }

    let count = 0;
    selectedPanels.forEach((panel) => {
      const changed = clearContentFromPanel(panel);
      if (changed) count++;
      updateAulasIndicators(panel);
    });

    statusBtn.textContent = `Limpo (${count} itens)!`;
    setTimeout(() => {
      statusBtn.disabled = false;
      statusBtn.textContent = "Limpar Conteúdos";
    }, 2000);
  }

  function insertMasterTools() {
    if (document.getElementById("poMasterToolsContainer")) return;
    const firstPanel = document.querySelector("mat-expansion-panel.mat-expansion-panel");
    if (!firstPanel || !firstPanel.parentElement) return;

    const container = document.createElement("div");
    container.id = "poMasterToolsContainer";

    container.innerHTML = `
      <label class="po-master-label" title="Marca/Desmarca tudo">
        <input type="checkbox" id="poMasterCheckbox" />
        <span class="po-master-text">Selecionar Todos</span>
      </label>

      <div class="po-separator-vert"></div>

      <label class="po-master-label" title="Marca/Desmarca apenas as primeiras aulas do dia">
        <input type="checkbox" id="poSelectFirstCheckbox" />
        <span class="po-master-text">Selecionar Primeiras Aulas</span>
      </label>

      <div class="po-separator-vert"></div>

      <button id="poCascadeBtn" class="po-master-btn po-btn-orange" title="Copia texto da 1ª aula do dia para as demais aulas DO MESMO DIA.">
        Replicar em Cascata (Dia)
      </button>

      <button id="poClearBtn" class="po-master-btn po-btn-red" title="Apaga texto de Previsto e Realizado dos selecionados.">
        Limpar Conteúdos
      </button>
    `;

    container.querySelector("#poMasterCheckbox").addEventListener("change", (e) => {
      const check = e.target.checked;
      document.querySelectorAll("mat-expansion-panel-header po-checkbox").forEach((cb) => {
        const ariaEl = cb.querySelector('[role="checkbox"]');
        if ((ariaEl && ariaEl.getAttribute("aria-checked") === "true") !== check) {
          (cb.querySelector(".po-checkbox-outline") || cb.querySelector("input") || cb).click();
        }
      });
    });

    container.querySelector("#poSelectFirstCheckbox").addEventListener("change", (e) => {
      const check = e.target.checked;
      const panels = document.querySelectorAll("mat-expansion-panel.mat-expansion-panel");
      let lastDate = null;
      panels.forEach((p) => {
        const h = p.querySelector("mat-expansion-panel-header");
        if (!h) return;
        const cb = h.querySelector("po-checkbox");
        if (!cb) return;
        const m = (h.innerText || "").match(REGEX_DATE);
        if (m) {
          if (m[1] !== lastDate) {
            lastDate = m[1];
            const ariaEl = cb.querySelector('[role="checkbox"]');
            if ((ariaEl && ariaEl.getAttribute("aria-checked") === "true") !== check) {
              (cb.querySelector(".po-checkbox-outline") || cb.querySelector("input") || cb).click();
            }
          }
        }
      });
    });

    const btnCascade = container.querySelector("#poCascadeBtn");
    btnCascade.addEventListener("click", (e) => {
      e.preventDefault();
      runCascadeCopy(btnCascade);
    });

    const btnClear = container.querySelector("#poClearBtn");
    btnClear.addEventListener("click", (e) => {
      e.preventDefault();
      if (
        confirm(
          "ATENÇÃO: Isso apagará o texto 'Previsto' e 'Realizado' de todas as aulas SELECIONADAS.\n\nDeseja continuar?"
        )
      ) {
        runClearContent(btnClear);
      }
    });

    firstPanel.parentElement.insertBefore(container, firstPanel);
  }

  // =======================
  // INIT
  // =======================
  function init() {
    if (document.body) ensurePanel();

    startSmartAutoExpand();

    applyFont(loadFontPx());
    installHeaderControls();
    applyLockState(loadLock());

    installAulasDelegatedListenersOnce();
    installAulasIndicators();
    isolateCheckboxesInHeaders();
    highlightByDateChange();
    insertMasterTools();

    renderHiddenList();
    runArchivedCleanup();

    // ✅ notas
    __poNotasPanelAutoToggle();
  }
  init();

  let scheduled = false;
  const obs = new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;

    runArchivedCleanup();

    requestAnimationFrame(() => {
      scheduled = false;

      if (document.body) ensurePanel();
      applyFont(loadFontPx());
      installHeaderControls();
      applyLockState(loadLock());

      const panel = document.getElementById("poFontPanel");
      if (panel && !panel.classList.contains("minimized") && !__poPanelIsResizing) {
        applyPanelSize(panel, loadPanelSize());
      }

      if (document.querySelector("mat-expansion-panel")) {
        installAulasIndicators();
        isolateCheckboxesInHeaders();
        highlightByDateChange();
        insertMasterTools();
      }

      // ✅ notas (painel separado + launcher)
      __poNotasPanelAutoToggle();
    });
  });

  obs.observe(document.documentElement, { childList: true, subtree: true });
})();