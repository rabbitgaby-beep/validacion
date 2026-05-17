// ══════════════════════════════════════════
        // ESTADO GLOBAL
        // ══════════════════════════════════════════
        const state = {
            sheets: {},          // { prog: [], prest: [], ind: [] } — arrays de col names
            data: {},            // { prog: [[...]], prest: [[...]], ind: [[...]] }
            headers: {},         // { prog: [...], prest: [...], ind: [...] }
            activeTab: 'prog',
            sections: [],        // array de section objects
            rightTab: 'stats',
            fkProgPrest: '',     // FK entre programas y prestaciones
            fkProgInd: '',       // FK entre programas e indicadores
            fkPrestInd: '',      // FK entre prestaciones e indicadores
            filterCol: '',       // columna para filtrar (ej: Ministerio)
            pdfConfig: {
                title: 'Informe de Programas Sociales',
                groupBy: '',       // agrupar PDFs por esta columna
                filterVigor: true, // solo Vigente = 1
                vigorCol: '',
            }
        };

        let sectionCounter = 0;
        let draggedField = null;
        let draggedRow = null;
        let activeOrg = null;

        // ══════════════════════════════════════════
        // CARGA DE EXCEL
        // ══════════════════════════════════════════
        const uploadZone = document.getElementById('upload-zone');
        const fileInput = document.getElementById('file-input');

        uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('dragover'); });
        uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
        uploadZone.addEventListener('drop', e => {
            e.preventDefault(); uploadZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) loadExcel(file);
        });
        fileInput.addEventListener('change', e => { if (e.target.files[0]) loadExcel(e.target.files[0]); });

        function loadExcel(file) {
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const wb = XLSX.read(e.target.result, { type: 'array' });
                    // Intentar detectar hojas automáticamente
                    const sheetMap = detectSheets(wb.SheetNames);

                    ['prog', 'prest', 'ind'].forEach(key => {
                        const sname = sheetMap[key];
                        if (!sname) return;
                        const ws = wb.Sheets[sname];
                        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                        if (json.length < 1) return;
                        state.headers[key] = json[0].map(String);
                        state.data[key] = json.slice(1);
                    });

                    document.getElementById('upload-zone').style.display = 'none';
                    document.getElementById('fields-container').style.display = 'flex';
                    document.getElementById('template-empty').style.display = 'none';
                    document.getElementById('btn-generate').disabled = false;

                    // Auto-detectar FKs
                    autoDetectFKs();

                    switchTab('prog');
                    updateRightPanel();
                    updateStepPills(2);
                } catch (err) {
                    alert('Error al leer el archivo: ' + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function detectSheets(names) {
            // Busca hojas que contengan las palabras clave
            const lower = names.map(n => n.toLowerCase());
            const find = kw => {
                const i = lower.findIndex(n => n.includes(kw));
                return i >= 0 ? names[i] : null;
            };
            return {
                prog: find('program') || find('prog') || names[0] || null,
                prest: find('prestac') || names[1] || null,
                ind: find('indicad') || names[2] || null,
            };
        }

        function autoDetectFKs() {
            // Buscar columnas tipo "Id" comunes entre hojas
            const ph = state.headers.prog || [];
            const prh = state.headers.prest || [];
            const ih = state.headers.ind || [];

            const commonProg = prh.find(c => ph.includes(c) && c.toLowerCase().includes('id')) ||
                prh.find(c => ph.includes(c));
            if (commonProg) state.fkProgPrest = commonProg;

            const commonInd = ih.find(c => ph.includes(c) && c.toLowerCase().includes('id')) ||
                ih.find(c => ph.includes(c));
            if (commonInd) state.fkProgInd = commonInd;

            const commonPrestInd = ih.find(c => prh.includes(c) && !ph.includes(c)) ||
                ih.find(c => prh.includes(c) && c !== commonInd);
            if (commonPrestInd) state.fkPrestInd = commonPrestInd;

            // Detectar columna de vigencia
            const vigorCol = ph.find(c => c.toLowerCase().includes('vigente'));
            if (vigorCol) state.pdfConfig.vigorCol = vigorCol;

            // Detectar columna para agrupación por organismo
            const groupCol = ph.find(c => c.toLowerCase().includes('ministerio') || c.toLowerCase().includes('organismo'));
            if (groupCol) state.pdfConfig.groupBy = groupCol;
        }

        // ══════════════════════════════════════════
        // PANEL DE CAMPOS
        // ══════════════════════════════════════════
        function switchTab(tab) {
            state.activeTab = tab;
            ['prog', 'prest', 'ind'].forEach(t => {
                const el = document.getElementById('tab-' + t);
                el.className = 'sheet-tab';
            });
            const colors = { prog: 'active-prog', prest: 'active-prest', ind: 'active-ind' };
            document.getElementById('tab-' + tab).classList.add(colors[tab]);
            renderFieldList();
        }

        function renderFieldList() {
            const list = document.getElementById('fields-list');
            const headers = state.headers[state.activeTab] || [];
            const usedCols = getUsedColumns();
            const colors = { prog: 'dot-prog', prest: 'dot-prest', ind: 'dot-ind' };
            const dotColor = colors[state.activeTab];

            list.innerHTML = headers.map((col, i) => {
                const used = usedCols.has(state.activeTab + '::' + col);
                return `<div class="field-chip ${used ? 'used' : ''}" 
              draggable="true"
              data-sheet="${state.activeTab}" data-col="${col}"
              ondragstart="onFieldDragStart(event, '${state.activeTab}', '${escAttr(col)}')"
              ondragend="onFieldDragEnd(event)"
              title="${col}">
      <span class="chip-dot ${dotColor}"></span>
      <span class="chip-name">${col}</span>
      <span class="chip-drag">⣿</span>
    </div>`;
            }).join('');
        }

        function getUsedColumns() {
            const used = new Set();
            state.sections.forEach(sec => {
                sec.rows.forEach(row => {
                    used.add(row.sheet + '::' + row.col);
                });
            });
            return used;
        }

        // ══════════════════════════════════════════
        // DRAG & DROP — CAMPOS A SECCIÓN
        // ══════════════════════════════════════════
        function onFieldDragStart(e, sheet, col) {
            draggedField = { sheet, col };
            draggedRow = null;
            e.dataTransfer.effectAllowed = 'copy';
            e.dataTransfer.setData('application/json', JSON.stringify({ sheet, col }));
            e.dataTransfer.setData('text/plain', col); // Fallback
            setTimeout(() => e.target.classList.add('dragging'), 0);
        }
        function onFieldDragEnd(e) {
            e.target.classList.remove('dragging');
            draggedField = null;
        }

        function onDropZoneDragEnter(e, secId) {
            e.preventDefault();
            document.getElementById('dz-' + secId).classList.add('drag-over');
            e.dataTransfer.dropEffect = 'move';
        }
        function onDropZoneDragOver(e, secId) {
            e.preventDefault();
            document.getElementById('dz-' + secId).classList.add('drag-over');
        }
        function onDropZoneDragLeave(secId) {
            document.getElementById('dz-' + secId).classList.remove('drag-over');
        }
        function onDropZoneDrop(e, secId) {
            e.preventDefault();
            e.stopPropagation();
            document.getElementById('dz-' + secId).classList.remove('drag-over');
            
            let sheet = null;
            let col = null;

            try {
                const data = JSON.parse(e.dataTransfer.getData('application/json'));
                if (data.type === 'reorder-sec') {
                    const sourceIdx = data.idx;
                    const targetIdx = state.sections.findIndex(s => s.id === secId);
                    if (sourceIdx !== -1 && targetIdx !== -1 && sourceIdx !== targetIdx) {
                        const movedSec = state.sections.splice(sourceIdx, 1)[0];
                        state.sections.splice(targetIdx, 0, movedSec);
                        renderTemplate();
                    }
                    return;
                }
                if (data && data.sheet && data.col) {
                    sheet = data.sheet;
                    col = data.col;
                }
            } catch (err) {
                console.warn('No se pudo parsear el dataTransfer json:', err);
            }

            if (!sheet && draggedField) {
                sheet = draggedField.sheet;
                col = draggedField.col;
            }

            if (sheet && col) {
                addRowToSection(secId, sheet, col);
            }
            draggedField = null;
        }

        // --- DRAG & DROP PARA SECCIONES ---
        function onSecDragStart(e, secIdx) {
            e.stopPropagation();
            draggedField = null;
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('application/json', JSON.stringify({ type: 'reorder-sec', idx: secIdx }));
            e.dataTransfer.setData('text/plain', 'sec-' + secIdx);
            setTimeout(() => e.target.classList.add('dragging-sec'), 0);
        }
        function onSecDragEnd(e) {
            e.stopPropagation();
            e.target.classList.remove('dragging-sec');
        }
        function onSecDragEnter(e) {
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'move';
        }
        function onSecDragOver(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        function onSecDrop(e, targetIdx) {
            e.preventDefault();
            e.stopPropagation();
            try {
                const data = JSON.parse(e.dataTransfer.getData('application/json'));
                if (data.type === 'reorder-sec') {
                    const sourceIdx = data.idx;
                    if (sourceIdx !== -1 && targetIdx !== -1 && sourceIdx !== targetIdx) {
                        const movedSec = state.sections.splice(sourceIdx, 1)[0];
                        state.sections.splice(targetIdx, 0, movedSec);
                        renderTemplate();
                    }
                }
            } catch (err) {}
        }

        function onRowDragStart(e, secId, idx) {
            e.stopPropagation();
            draggedField = null;
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('application/json', JSON.stringify({ type: 'reorder', secId, idx }));
            e.dataTransfer.setData('text/plain', 'row-' + idx);
            setTimeout(() => e.target.classList.add('dragging'), 0);
        }
        function onRowDragEnd(e) {
            e.stopPropagation();
            e.target.classList.remove('dragging');
        }
        function onRowDragEnter(e) {
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'move';
        }
        function onRowDragOver(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        function onRowDrop(e, targetSecId, targetIdx) {
            e.preventDefault();
            e.stopPropagation();
            
            try {
                const data = JSON.parse(e.dataTransfer.getData('application/json'));
                if (data.type === 'reorder-sec') {
                    const sourceIdx = data.idx;
                    const destIdx = state.sections.findIndex(s => s.id === targetSecId);
                    if (sourceIdx !== -1 && destIdx !== -1 && sourceIdx !== destIdx) {
                        const movedSec = state.sections.splice(sourceIdx, 1)[0];
                        state.sections.splice(destIdx, 0, movedSec);
                        renderTemplate();
                    }
                    return;
                }
                if (data.type === 'reorder') {
                    const sec = state.sections.find(s => s.id === data.secId);
                    const targetSec = state.sections.find(s => s.id === targetSecId);
                    if (sec && targetSec) {
                        const [movedRow] = sec.rows.splice(data.idx, 1);
                        targetSec.rows.splice(targetIdx, 0, movedRow);
                        renderTemplate();
                    }
                } else if (data.sheet && data.col) {
                    const sec = state.sections.find(s => s.id === targetSecId);
                    if (sec && !sec.rows.find(r => r.sheet === data.sheet && r.col === data.col)) {
                        sec.rows.splice(targetIdx, 0, {
                            sheet: data.sheet, col: data.col, label: data.col,
                            bold: false, big: false, bullet: false, table: false
                        });
                        renderTemplate();
                        renderFieldList();
                    }
                } else { throw new Error('fallback'); }
            } catch (err) {
                if (draggedField) {
                    const sec = state.sections.find(s => s.id === targetSecId);
                    if (sec && !sec.rows.find(r => r.sheet === draggedField.sheet && r.col === draggedField.col)) {
                        sec.rows.splice(targetIdx, 0, {
                            sheet: draggedField.sheet, col: draggedField.col, label: draggedField.col,
                            bold: false, big: false, bullet: false, table: false
                        });
                        renderTemplate();
                        renderFieldList();
                    }
                }
            }
            draggedField = null;
        }

        // ══════════════════════════════════════════
        // SECCIONES Y FILAS
        // ══════════════════════════════════════════
        const SECTION_COLORS = {
            prog: '#6ab0ff', prest: '#2dd4a0', ind: '#b06aff', mixed: '#f5a623'
        };

        function addSection(title, type) {
            if (!state.headers.prog) { alert('Primero cargá un Excel.'); return; }
            const id = 's' + (++sectionCounter);
            const sec = { id, title, type, rows: [] };
            state.sections.push(sec);
            renderTemplate();
        }

        function deleteSection(secId) {
            state.sections = state.sections.filter(s => s.id !== secId);
            renderTemplate();
            renderFieldList();
        }

        function getRowId(r) { return r.id || (r.sheet + '::' + r.col); }

        function addRowToSection(secId, sheet, col) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            if (sec.rows.find(r => r.sheet === sheet && r.col === col && !r.type)) return;
            sec.rows.push({
                sheet, col,
                label: col,
                bold: false, big: false, bullet: false, table: false
            });
            renderTemplate();
            renderFieldList();
        }

        function addSpecialRow(secId, type) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            const uniqueId = Math.random().toString(36).substr(2, 9);
            if (type === 'text-free') {
                sec.rows.push({
                    type: 'text-free',
                    id: uniqueId,
                    value: '',
                    bold: false, big: false, bullet: false, table: false
                });
            } else if (type === 'dropdown-map') {
                sec.rows.push({
                    type: 'dropdown-map',
                    id: uniqueId,
                    label: 'Campo Dinámico',
                    selectedSheet: '',
                    selectedCol: '',
                    bold: false, big: false, bullet: false, table: false
                });
            }
            renderTemplate();
        }

        function deleteRow(secId, rowId) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            sec.rows = sec.rows.filter(r => getRowId(r) !== rowId);
            renderTemplate();
            renderFieldList();
        }

        function moveSection(idx, dir) {
            const newIdx = idx + dir;
            if (newIdx < 0 || newIdx >= state.sections.length) return;
            const temp = state.sections[idx];
            state.sections[idx] = state.sections[newIdx];
            state.sections[newIdx] = temp;
            renderTemplate();
        }

        function toggleFmt(secId, rowId, fmt) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            const row = sec.rows.find(r => getRowId(r) === rowId);
            if (!row) return;
            row[fmt] = !row[fmt];
            renderTemplate();
        }

        function updateLabel(secId, rowId, val) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            const row = sec.rows.find(r => getRowId(r) === rowId);
            if (row) row.label = val;
        }

        function updateValue(secId, rowId, val) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            const row = sec.rows.find(r => getRowId(r) === rowId);
            if (row) row.value = val;
        }

        function updateDropdown(secId, rowId, val) {
            const sec = state.sections.find(s => s.id === secId);
            if (!sec) return;
            const row = sec.rows.find(r => getRowId(r) === rowId);
            if (row) {
                if (!val) {
                    row.selectedSheet = '';
                    row.selectedCol = '';
                } else {
                    const parts = val.split('::');
                    row.selectedSheet = parts[0];
                    row.selectedCol = parts[1];
                }
                renderTemplate();
            }
        }

        function updateSectionTitle(secId, val) {
            const sec = state.sections.find(s => s.id === secId);
            if (sec) sec.title = val;
        }

        // ══════════════════════════════════════════
        // RENDER PLANTILLA
        // ══════════════════════════════════════════
        function renderTemplate() {
            const canvas = document.getElementById('template-canvas');

            if (state.sections.length === 0) {
                canvas.innerHTML = `
                    <div class="empty-state" id="template-empty">
                        <div class="empty-icon">📋</div>
                        <strong style="color:var(--text2)">Cargá un Excel primero</strong><br>
                        Luego arrastrá campos desde la izquierda<br>hacia las secciones de la plantilla,<br>o usá los
                        botones de arriba para agregar secciones.
                    </div>`;
                return;
            }

            canvas.innerHTML = '';

            state.sections.forEach((sec, secIndex) => {
                const color = SECTION_COLORS[sec.type] || SECTION_COLORS.mixed;
                const colKey = c => `${c.sheet}::${c.col}`;

                // ¿Es sección de prestaciones o indicadores? Mostrar config FK
                let fkHtml = '';
                if (sec.type === 'prest') {
                    const opts = (state.headers.prest || []).map(c => `<option value="${escAttr(c)}" ${state.fkProgPrest === c ? 'selected' : ''}>${c}</option>`).join('');
                    fkHtml = `<div class="fk-config">
        <span class="fk-label">FK Prog→Prest:</span>
        <select class="fk-select" onchange="state.fkProgPrest=this.value">${opts}</select>
      </div>`;
                } else if (sec.type === 'ind') {
                    const optsP = (state.headers.ind || []).map(c => `<option value="${escAttr(c)}" ${state.fkProgInd === c ? 'selected' : ''}>${c}</option>`).join('');
                    const optsPr = (state.headers.ind || []).map(c => `<option value="${escAttr(c)}" ${state.fkPrestInd === c ? 'selected' : ''}>${c}</option>`).join('');
                    fkHtml = `<div class="fk-config">
        <span class="fk-label">FK→Prog:</span>
        <select class="fk-select" style="width:auto" onchange="state.fkProgInd=this.value">${optsP}</select>
        <span class="fk-label" style="margin-left:8px">FK→Prest:</span>
        <select class="fk-select" style="width:auto" onchange="state.fkPrestInd=this.value">${optsPr}</select>
      </div>`;
                }

                const rowsHtml = sec.rows.map((r, rowIndex) => {
                    const ck = getRowId(r);
                    const dotClass = r.type ? 'dot-special' : ({ prog: 'dot-prog', prest: 'dot-prest', ind: 'dot-ind' }[r.sheet] || '');
                    const fmtBtn = (fmt, icon) => `<button class="fmt-btn ${r[fmt] ? 'active' : ''}" onclick="toggleFmt('${sec.id}','${ck}','${fmt}')" title="${fmt}">${icon}</button>`;
                    
                    let innerHtml = '';
                    if (r.type === 'text-free') {
                        innerHtml = `<input class="row-label-input" type="text" value="${escAttr(r.value)}" placeholder="Escribí texto libre..." 
                          onchange="updateValue('${sec.id}','${ck}',this.value)"
                          oninput="updateValue('${sec.id}','${ck}',this.value)" style="flex:1;">
                        <span class="row-field-name" style="width:auto">📝</span>`;
                    } else if (r.type === 'dropdown-map') {
                        let opts = '<option value="">Seleccionar columna...</option>';
                        ['prog', 'prest', 'ind'].forEach(s => {
                            if (state.headers[s]) {
                                state.headers[s].forEach(c => {
                                    const val = `${s}::${c}`;
                                    const sel = (r.selectedSheet === s && r.selectedCol === c) ? 'selected' : '';
                                    opts += `<option value="${escAttr(val)}" ${sel}>[${s.toUpperCase()}] ${escAttr(c)}</option>`;
                                });
                            }
                        });
                        innerHtml = `<input class="row-label-input" type="text" value="${escAttr(r.label)}" placeholder="Label..." 
                          onchange="updateLabel('${sec.id}','${ck}',this.value)"
                          oninput="updateLabel('${sec.id}','${ck}',this.value)" style="width: 120px;">
                        <select class="row-select-input" onchange="updateDropdown('${sec.id}','${ck}',this.value)" style="flex:1; margin:0 4px; border:1px solid var(--border); border-radius:4px; padding:2px;">${opts}</select>
                        <span class="row-field-name" style="width:auto">🔄</span>`;
                    } else {
                        innerHtml = `<input class="row-label-input" type="text" value="${escAttr(r.label)}" placeholder="Label..." 
                          onchange="updateLabel('${sec.id}','${ck}',this.value)"
                          oninput="updateLabel('${sec.id}','${ck}',this.value)">
                        <span class="row-field-name">{${r.col}}</span>`;
                    }

                    return `<div class="template-row" draggable="true"
          ondragstart="onRowDragStart(event, '${sec.id}', ${rowIndex})"
          ondragend="onRowDragEnd(event)"
          ondragenter="onRowDragEnter(event)"
          ondragover="onRowDragOver(event)"
          ondrop="onRowDrop(event, '${sec.id}', ${rowIndex})">
        <span class="row-handle">⣿</span>
        <span class="row-dot ${dotClass}"></span>
        ${innerHtml}
        <div class="row-format-btns">
          ${fmtBtn('bold', 'B')}
          ${fmtBtn('big', 'A+')}
          ${fmtBtn('bullet', '•')}
          ${fmtBtn('table', '⊞')}
        </div>
        <button class="row-del" onclick="deleteRow('${sec.id}','${ck}')">✕</button>
      </div>`;
                }).join('');

                const dropHint = sec.rows.length === 0
                    ? '<div class="drop-zone-hint">↙ Arrastrá campos aquí</div>'
                    : '';

                const el = document.createElement('div');
                el.className = 'template-section';
                el.ondragstart = (e) => onSecDragStart(e, secIndex);
                el.ondragend = (e) => onSecDragEnd(e);
                el.ondragenter = (e) => onSecDragEnter(e);
                el.ondragover = (e) => onSecDragOver(e);
                el.ondrop = (e) => onSecDrop(e, secIndex);

                const upDisabled = secIndex === 0 ? 'disabled' : '';
                const downDisabled = secIndex === state.sections.length - 1 ? 'disabled' : '';

                el.innerHTML = `
      <div class="section-header">
        <span class="section-drag-handle" title="Arrastrar sección"
              onmousedown="this.closest('.template-section').draggable = true;"
              onmouseup="this.closest('.template-section').draggable = false;"
              onmouseleave="this.closest('.template-section').draggable = false;">⣿</span>
        <div class="sec-move-btns">
            <button class="sec-btn-move" aria-label="Mover arriba" onclick="moveSection(${secIndex}, -1)" ${upDisabled}>↑</button>
            <button class="sec-btn-move" aria-label="Mover abajo" onclick="moveSection(${secIndex}, 1)" ${downDisabled}>↓</button>
        </div>
        <span class="section-color" style="background:${color}; border-radius:2px"></span>
        <input class="section-title-input" type="text" value="${escAttr(sec.title)}"
          onchange="updateSectionTitle('${sec.id}',this.value)"
          oninput="updateSectionTitle('${sec.id}',this.value)">
        <span class="badge badge-${sec.type}">${sec.type}</span>
        <button class="section-del" onclick="deleteSection('${sec.id}')">✕</button>
      </div>
      <div class="section-body">
        ${fkHtml}
        <div class="drop-zone" id="dz-${sec.id}"
          ondragenter="onDropZoneDragEnter(event,'${sec.id}')"
          ondragover="onDropZoneDragOver(event,'${sec.id}')"
          ondragleave="onDropZoneDragLeave('${sec.id}')"
          ondrop="onDropZoneDrop(event,'${sec.id}')">
          ${rowsHtml}
          ${dropHint}
        </div>
        <div class="special-fields-btns" style="padding: 0 12px 10px; display: flex; gap: 8px;">
            <button class="action-btn btn-add-special" onclick="addSpecialRow('${sec.id}', 'text-free')">+ Texto Libre</button>
            <button class="action-btn btn-add-special" onclick="addSpecialRow('${sec.id}', 'dropdown-map')">+ Campo Dinámico</button>
        </div>
      </div>`;
                canvas.appendChild(el);
            });
        }

        // ══════════════════════════════════════════
        // PANEL DERECHO
        // ══════════════════════════════════════════
        function switchRightTab(tab) {
            state.rightTab = tab;
            ['stats', 'config', 'preview'].forEach(t => {
                document.getElementById('rtab-' + t).className = 'preview-tab' + (t === tab ? ' active' : '');
            });
            updateRightPanel();
        }

        function updateRightPanel() {
            const body = document.getElementById('right-panel-body');
            if (state.rightTab === 'stats') renderStatsPanel(body);
            else if (state.rightTab === 'config') renderConfigPanel(body);
            else renderPreviewPanel(body);
        }

        function renderStatsPanel(body) {
            if (!state.headers.prog) {
                body.innerHTML = '<div class="empty-state"><div class="empty-icon">📂</div>Cargá un Excel primero</div>';
                return;
            }
            const progRows = state.data.prog || [];
            const prestRows = state.data.prest || [];
            const indRows = state.data.ind || [];

            const vigentCol = state.headers.prog.findIndex(c => c === state.pdfConfig.vigorCol);
            const vigentes = vigentCol >= 0 ? progRows.filter(r => String(r[vigentCol]) === '1').length : progRows.length;

            const groupCol = state.headers.prog.findIndex(c => c === state.pdfConfig.groupBy);
            const orgs = groupCol >= 0
                ? [...new Set(progRows.map(r => r[groupCol]).filter(Boolean))].length
                : 0;

            body.innerHTML = `
    <div class="config-section">
      <div class="config-title">Estadísticas del Excel</div>
      <div class="stat-row"><span class="stat-key">Programas totales</span><span class="stat-val">${progRows.length}</span></div>
      <div class="stat-row"><span class="stat-key">Programas vigentes</span><span class="stat-val green">${vigentes}</span></div>
      <div class="stat-row"><span class="stat-key">Prestaciones</span><span class="stat-val">${prestRows.length}</span></div>
      <div class="stat-row"><span class="stat-key">Indicadores</span><span class="stat-val">${indRows.length}</span></div>
      ${orgs ? `<div class="stat-row"><span class="stat-key">Organismos</span><span class="stat-val">${orgs}</span></div>` : ''}
    </div>
    <div class="config-section">
      <div class="config-title">Columnas detectadas</div>
      <div class="stat-row"><span class="stat-key">Prog. → Prest. (FK)</span><span class="stat-val" style="font-size:9px;max-width:120px;overflow:hidden;text-overflow:ellipsis">${state.fkProgPrest || '—'}</span></div>
      <div class="stat-row"><span class="stat-key">Prog. → Ind. (FK)</span><span class="stat-val" style="font-size:9px;max-width:120px;overflow:hidden;text-overflow:ellipsis">${state.fkProgInd || '—'}</span></div>
      <div class="stat-row"><span class="stat-key">Prest. → Ind. (FK)</span><span class="stat-val" style="font-size:9px;max-width:120px;overflow:hidden;text-overflow:ellipsis">${state.fkPrestInd || '—'}</span></div>
      <div class="stat-row"><span class="stat-key">Columna Vigente</span><span class="stat-val" style="font-size:9px">${state.pdfConfig.vigorCol || '—'}</span></div>
      <div class="stat-row"><span class="stat-key">Agrupar por</span><span class="stat-val" style="font-size:9px;max-width:120px;overflow:hidden;text-overflow:ellipsis">${state.pdfConfig.groupBy || '—'}</span></div>
    </div>
    <div class="config-section">
      <div class="config-title">Plantilla</div>
      <div class="stat-row"><span class="stat-key">Secciones</span><span class="stat-val">${state.sections.length}</span></div>
      <div class="stat-row"><span class="stat-key">Campos mapeados</span><span class="stat-val green">${state.sections.reduce((s, sec) => s + sec.rows.length, 0)}</span></div>
    </div>`;
        }

        function renderConfigPanel(body) {
            const progH = state.headers.prog || [];
            const makeOpts = (selected) => progH.map(c => `<option value="${escAttr(c)}" ${selected === c ? 'selected' : ''}>${c}</option>`).join('');

            body.innerHTML = `
    <div class="config-section">
      <div class="config-title">Configuración del PDF</div>
      <div class="config-row"><div>
        <div style="font-size:10px;color:var(--text3);margin-bottom:4px">Título del informe</div>
        <input class="config-input" type="text" value="${escAttr(state.pdfConfig.title)}"
          oninput="state.pdfConfig.title=this.value">
      </div></div>
    </div>
    <div class="config-section">
      <div class="config-title">Filtros</div>
      <div class="config-row">
        <input type="checkbox" class="config-checkbox" id="chk-vigor" ${state.pdfConfig.filterVigor ? 'checked' : ''}
          onchange="state.pdfConfig.filterVigor=this.checked">
        <label for="chk-vigor" class="config-label">Solo programas vigentes</label>
      </div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:4px">Columna "Vigente"</div>
      <select class="config-input config-select" onchange="state.pdfConfig.vigorCol=this.value">
        <option value="">— ninguna —</option>
        ${makeOpts(state.pdfConfig.vigorCol)}
      </select>
    </div>
    <div class="config-section">
      <div class="config-title">Agrupación de PDFs</div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:4px">Generar un PDF por valor de:</div>
      <select class="config-input config-select" onchange="state.pdfConfig.groupBy=this.value">
        <option value="">— un solo PDF —</option>
        ${makeOpts(state.pdfConfig.groupBy)}
      </select>
      <div style="font-size:10px;color:var(--text3);margin-top:8px">
        Si seleccionás "Ministerio", se genera un PDF separado por cada ministerio.
      </div>
    </div>
    <div class="config-section">
      <div class="config-title">FKs / Relaciones</div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:4px">Prog → Prest (FK)</div>
      <select class="config-input config-select" onchange="state.fkProgPrest=this.value">
        ${(state.headers.prest || []).map(c => `<option value="${escAttr(c)}" ${state.fkProgPrest === c ? 'selected' : ''}>${c}</option>`).join('')}
      </select>
      <div style="font-size:10px;color:var(--text3);margin:8px 0 4px">Prog → Ind (FK)</div>
      <select class="config-input config-select" onchange="state.fkProgInd=this.value">
        ${(state.headers.ind || []).map(c => `<option value="${escAttr(c)}" ${state.fkProgInd === c ? 'selected' : ''}>${c}</option>`).join('')}
      </select>
      <div style="font-size:10px;color:var(--text3);margin:8px 0 4px">Prest → Ind (FK)</div>
      <select class="config-input config-select" onchange="state.fkPrestInd=this.value">
        ${(state.headers.ind || []).map(c => `<option value="${escAttr(c)}" ${state.fkPrestInd === c ? 'selected' : ''}>${c}</option>`).join('')}
      </select>
    </div>`;
        }

        function renderPreviewPanel(body) {
            if (!state.headers.prog || state.sections.length === 0) {
                body.innerHTML = '<div class="empty-state"><div class="empty-icon">📄</div>Armá la plantilla primero</div>';
                return;
            }
            const progRows = state.data.prog || [];
            if (progRows.length === 0) { body.innerHTML = '<div class="empty-state">Sin datos</div>'; return; }
            const html = buildProgramHTML(progRows[0], state.headers.prog, 0);
            body.innerHTML = `<div class="pdf-preview-card" style="font-size:9px">${html}</div>`;
        }

        // ══════════════════════════════════════════
        // CONSTRUCCIÓN DE HTML DEL PDF
        // ══════════════════════════════════════════
        function getVal(row, headers, col) {
            const i = headers.indexOf(col);
            if (i < 0) return '';
            const v = row[i];
            if (v === null || v === undefined || v === '') return '';
            return String(v).trim();
        }

        function buildProgramHTML(progRow, progHeaders, progIdx) {
            const progName = getVal(progRow, progHeaders, 'Nombre del programa ingresado') ||
                getVal(progRow, progHeaders, 'Nombre del programa precargado') ||
                `Programa ${progIdx + 1}`;

            // Valor del FK del programa actual
            const progFKVal = state.fkProgPrest ? getVal(progRow, progHeaders, state.fkProgPrest) : null;
            // También buscar en indicadores via fkProgInd
            const progFKIndVal = state.fkProgInd ? getVal(progRow, progHeaders, state.fkProgInd) : null;

            let html = `<h1>${progName}</h1>`;

            state.sections.forEach(sec => {
                // Sección de programa
                if (sec.type === 'prog') {
                    html += `<div class="pdf-sec">${sec.title}</div>`;
                    sec.rows.forEach(r => {
                        if (r.sheet && r.sheet !== 'prog' && !r.type) return;
                        const val = getRowValue(r, progRow, null, null);
                        if (!val && r.type !== 'text-free') return;
                        html += buildFieldHTML(r, val);
                    });
                }
                // Sección de prestaciones
                else if (sec.type === 'prest') {
                    const prestRows = state.data.prest || [];
                    const prestH = state.headers.prest || [];
                    const fkCol = state.fkProgPrest;
                    const relPrest = fkCol && progFKVal
                        ? prestRows.filter(pr => getVal(pr, prestH, fkCol) === progFKVal)
                        : prestRows;

                    if (relPrest.length === 0) return;
                    html += `<div class="pdf-sec">${sec.title}</div>`;

                    relPrest.forEach((pr, pi) => {
                        const prestName = getVal(pr, prestH, 'Prestación') || `Prestación ${pi + 1}`;
                        html += `<div style="margin-bottom:6px;padding:5px 8px;border:1px solid #003366;border-radius:3px">`;
                        html += `<div style="font-weight:700;color:#003366;font-size:10px;margin-bottom:4px">Prestación ${pi + 1}: ${prestName}</div>`;
                        sec.rows.forEach(r => {
                            if (r.sheet && r.sheet !== 'prest' && !r.type) return;
                            const val = getRowValue(r, progRow, pr, null);
                            if (!val && r.type !== 'text-free') return;
                            html += buildFieldHTML(r, val);
                        });

                        // Indicadores de esta prestación
                        const indSec = state.sections.find(s => s.type === 'ind');
                        if (indSec && indSec.rows.length > 0) {
                            const indH = state.headers.ind || [];
                            const indRows = state.data.ind || [];
                            const prFKVal = state.fkPrestInd ? getVal(pr, prestH, state.fkPrestInd) : null;
                            const progFKInds = state.fkProgInd ? getVal(pr, prestH, state.fkProgInd) : null;

                            let relInds = indRows;
                            if (prFKVal && state.fkPrestInd) {
                                const fkIndCol = state.fkPrestInd;
                                relInds = indRows.filter(ir => getVal(ir, indH, fkIndCol) === prFKVal);
                            } else if (progFKIndVal && state.fkProgInd) {
                                relInds = indRows.filter(ir => getVal(ir, indH, state.fkProgInd) === progFKIndVal);
                            }

                            if (relInds.length > 0) {
                                // Buscar si algún row tiene flag table
                                const hasTable = indSec.rows.some(r => r.table);
                                if (hasTable) {
                                    const tableCols = indSec.rows.filter(r => (r.sheet === 'ind' || r.type) && r.table);
                                    html += `<div style="margin-top:6px"><div style="font-size:9px;font-weight:700;margin-bottom:3px">Indicadores</div>`;
                                    html += `<table><thead><tr>${tableCols.map(r => `<th>${r.label}</th>`).join('')}</tr></thead><tbody>`;
                                    relInds.forEach(ir => {
                                        html += `<tr>${tableCols.map(r => `<td>${getRowValue(r, progRow, pr, ir) || 'S/I'}</td>`).join('')}</tr>`;
                                    });
                                    html += `</tbody></table></div>`;
                                } else {
                                    html += `<div style="margin-top:4px;font-size:9px;font-weight:700">Indicadores:</div>`;
                                    relInds.forEach(ir => {
                                        indSec.rows.forEach(r => {
                                            if (r.sheet && r.sheet !== 'ind' && !r.type) return;
                                            const val = getRowValue(r, progRow, pr, ir);
                                            if (!val && r.type !== 'text-free') return;
                                            html += buildFieldHTML(r, val);
                                        });
                                    });
                                }
                            }
                        }
                        html += `</div>`;
                    });
                }
                // Sección indicadores standalone (solo si no hay prestaciones)
                else if (sec.type === 'ind') {
                    const hasPrestSec = state.sections.some(s => s.type === 'prest');
                    if (hasPrestSec) return; // ya se renderizaron dentro de prestaciones
                    const indH = state.headers.ind || [];
                    const indRows = state.data.ind || [];
                    const relInds = state.fkProgInd && progFKIndVal
                        ? indRows.filter(ir => getVal(ir, indH, state.fkProgInd) === progFKIndVal)
                        : indRows;
                    if (relInds.length === 0) return;
                    html += `<div class="pdf-sec">${sec.title}</div>`;
                    const hasTable = sec.rows.some(r => r.table);
                    const tableCols = sec.rows.filter(r => (r.sheet === 'ind' || r.type) && r.table);
                    if (hasTable && tableCols.length > 0) {
                        html += `<table><thead><tr>${tableCols.map(r => `<th>${r.label}</th>`).join('')}</tr></thead><tbody>`;
                        relInds.forEach(ir => {
                            html += `<tr>${tableCols.map(r => `<td>${getRowValue(r, progRow, null, ir) || 'S/I'}</td>`).join('')}</tr>`;
                        });
                        html += `</tbody></table>`;
                    } else {
                        relInds.forEach(ir => {
                            sec.rows.forEach(r => {
                                if (r.sheet && r.sheet !== 'ind' && !r.type) return;
                                const val = getRowValue(r, progRow, null, ir);
                                if (!val && r.type !== 'text-free') return;
                                html += buildFieldHTML(r, val);
                            });
                        });
                    }
                }
            });

            return html;
        }

        function buildFieldHTML(row, val) {
            const labelStr = row.label ? `<strong>${row.label}:</strong> ` : '';
            const style = [
                row.bold ? 'font-weight:700' : '',
                row.big ? 'font-size:12pt' : '',
            ].filter(Boolean).join(';');

            if (row.bullet) {
                const items = val.split(/[,;]\s*/).filter(Boolean);
                return `<div class="pdf-field"><div class="pdf-field-bold">${row.label}:</div><ul class="pdf-ul">${items.map(i => `<li>${i.trim()}</li>`).join('')}</ul></div>`;
            }
            return `<div class="pdf-field" ${style ? `style="${style}"` : ''}>${labelStr}${val}</div>`;
        }

        // ══════════════════════════════════════════
        // VISTA PREVIA (MODAL)
        // ══════════════════════════════════════════
        function showPreview() {
            if (!state.headers.prog || state.sections.length === 0) {
                alert('Cargá un Excel y armá la plantilla primero.');
                return;
            }
            const progRows = state.data.prog || [];
            if (progRows.length === 0) { alert('Sin datos.'); return; }

            const progH = state.headers.prog;
            const vigorCol = state.pdfConfig.vigorCol;
            const vigorIdx = vigorCol ? progH.indexOf(vigorCol) : -1;
            const rows = state.pdfConfig.filterVigor && vigorIdx >= 0
                ? progRows.filter(r => String(r[vigorIdx]) === '1')
                : progRows;

            const html = buildProgramHTML(rows[0] || progRows[0], progH, 0);
            document.getElementById('modal-body').innerHTML = `<div class="pdf-doc">${html}</div>`;
            document.getElementById('modal-overlay').classList.add('show');
        }
        function closeModal() {
            document.getElementById('modal-overlay').classList.remove('show');
        }

        // ══════════════════════════════════════════
        // GENERACIÓN MASIVA
        // ══════════════════════════════════════════
        function generateAll() {
            if (!state.headers.prog || state.sections.length === 0) {
                alert('Cargá el Excel y armá la plantilla primero.');
                return;
            }

            const progH = state.headers.prog;
            const vigorCol = state.pdfConfig.vigorCol;
            const vigorIdx = vigorCol ? progH.indexOf(vigorCol) : -1;
            let progRows = state.data.prog || [];

            if (state.pdfConfig.filterVigor && vigorIdx >= 0) {
                progRows = progRows.filter(r => String(r[vigorIdx]) === '1');
            }

            // Agrupar por organismo
            const groupCol = state.pdfConfig.groupBy;
            const groupIdx = groupCol ? progH.indexOf(groupCol) : -1;

            let groups = {};
            if (groupIdx >= 0) {
                progRows.forEach(r => {
                    const g = String(r[groupIdx] || 'Sin clasificar').trim();
                    if (!groups[g]) groups[g] = [];
                    groups[g].push(r);
                });
            } else {
                groups['Todos los Programas'] = progRows;
            }

            // Renderizar output page
            const orgSidebar = document.getElementById('org-sidebar');
            const pdfContainer = document.getElementById('pdf-container');

            orgSidebar.innerHTML = Object.keys(groups).map(g =>
                `<div class="org-item" onclick="scrollToOrg('${escAttr(g)}')" id="org-${escAttr(g).replace(/\W/g, '_')}">${g} <span class="org-count">(${groups[g].length})</span></div>`
            ).join('');

            pdfContainer.innerHTML = '';
            Object.keys(groups).forEach(g => {
                const anchor = document.createElement('div');
                anchor.id = 'anchor-' + g.replace(/\W/g, '_');

                const doc = document.createElement('div');
                doc.className = 'pdf-doc';

                let docHTML = `<h1>${state.pdfConfig.title}</h1>`;
                docHTML += `<div class="pdf-sec">${groupCol ? g : ''}</div>`;

                groups[g].forEach((row, i) => {
                    if (i > 0) docHTML += `<div class="prog-sep"></div>`;
                    docHTML += buildProgramHTML(row, progH, i);
                });

                doc.innerHTML = docHTML;
                pdfContainer.appendChild(anchor);
                pdfContainer.appendChild(doc);
            });

            document.getElementById('output-title').textContent = `${Object.keys(groups).length} reporte(s) generados — ${progRows.length} programas`;
            document.getElementById('output-page').classList.add('show');

            if (Object.keys(groups).length > 0) {
                scrollToOrg(Object.keys(groups)[0]);
            }

            updateStepPills(3);
        }

        function scrollToOrg(name) {
            activeOrg = name;
            const id = 'anchor-' + name.replace(/\W/g, '_');
            const el = document.getElementById(id);
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            document.querySelectorAll('.org-item').forEach(el => el.classList.remove('active'));
            const oid = 'org-' + name.replace(/\W/g, '_');
            const oEl = document.getElementById(oid);
            if (oEl) oEl.classList.add('active');
        }

        function closeOutput() {
            document.getElementById('output-page').classList.remove('show');
        }

        function printCurrent() {
            window.print();
        }

        async function fetchFont(url) {
            const res = await fetch(url);
            const buffer = await res.arrayBuffer();
            let binary = '';
            const bytes = new Uint8Array(buffer);
            const len = bytes.byteLength;
            for (let i = 0; i < len; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            return window.btoa(binary);
        }

        async function downloadPDF() {
            if (!state.headers.prog || state.sections.length === 0) {
                alert('Cargá el Excel y armá la plantilla primero.');
                return;
            }

            // Cambiamos el texto del botón temporalmente
            const btns = document.querySelectorAll('.btn-generate');
            const oldTexts = [];
            btns.forEach(btn => {
                oldTexts.push(btn.textContent);
                btn.textContent = "⏳ Cargando tipografía y generando PDF...";
                btn.disabled = true;
            });

            try {
                // Cargar fuentes Montserrat dinámicamente si no están en VFS
                if (!pdfMake.vfs || !pdfMake.vfs['Montserrat-Regular.ttf']) {
                    const regular = await fetchFont('https://raw.githubusercontent.com/JulietaUla/Montserrat/master/fonts/ttf/Montserrat-Regular.ttf');
                    const bold = await fetchFont('https://raw.githubusercontent.com/JulietaUla/Montserrat/master/fonts/ttf/Montserrat-Bold.ttf');
                    
                    pdfMake.vfs = pdfMake.vfs || {};
                    pdfMake.vfs['Montserrat-Regular.ttf'] = regular;
                    pdfMake.vfs['Montserrat-Bold.ttf'] = bold;
                    pdfMake.fonts = {
                        Roboto: {
                            normal: 'Roboto-Regular.ttf',
                            bold: 'Roboto-Medium.ttf',
                            italics: 'Roboto-Italic.ttf',
                            bolditalics: 'Roboto-MediumItalic.ttf'
                        },
                        Montserrat: {
                            normal: 'Montserrat-Regular.ttf',
                            bold: 'Montserrat-Bold.ttf',
                            italics: 'Montserrat-Regular.ttf',
                            bolditalics: 'Montserrat-Bold.ttf'
                        }
                    };
                }

                const headerImg = "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";

                const docDef = buildDocDefinition(headerImg);
                pdfMake.createPdf(docDef).download((state.pdfConfig.title || 'Reportes') + '.pdf');
            } catch (e) {
                console.error(e);
                alert('Error al generar PDF: ' + e.message);
            }

            btns.forEach((btn, i) => {
                btn.textContent = oldTexts[i];
                btn.disabled = false;
            });
        }

        const tableLayout = {
            hLineWidth: function (i, node) { return 1; },
            vLineWidth: function (i, node) { return 1; },
            hLineColor: function (i, node) { return '#dddddd'; },
            vLineColor: function (i, node) { return '#dddddd'; },
            paddingLeft: function (i, node) { return 8; },
            paddingRight: function (i, node) { return 8; },
            paddingTop: function (i, node) { return 5; },
            paddingBottom: function (i, node) { return 5; }
        };

        function getRowValue(r, progRow, pr, ir) {
            if (r.type === 'text-free') return r.value;
            
            const sheet = r.type === 'dropdown-map' ? r.selectedSheet : r.sheet;
            const col = r.type === 'dropdown-map' ? r.selectedCol : r.col;
            
            if (!sheet || !col) return null;
            
            if (sheet === 'prog') {
                return getVal(progRow, state.headers.prog || [], col);
            } else if (sheet === 'prest') {
                if (!pr) return null;
                return getVal(pr, state.headers.prest || [], col);
            } else if (sheet === 'ind') {
                if (!ir) return null;
                return getVal(ir, state.headers.ind || [], col);
            }
            return null;
        }

        function buildDocDefinition(headerImgBase64) {
            const progH = state.headers.prog;
            const vigorCol = state.pdfConfig.vigorCol;
            const vigorIdx = vigorCol ? progH.indexOf(vigorCol) : -1;
            let progRows = state.data.prog || [];

            if (state.pdfConfig.filterVigor && vigorIdx >= 0) {
                progRows = progRows.filter(r => String(r[vigorIdx]) === '1');
            }

            const groupCol = state.pdfConfig.groupBy;
            const groupIdx = groupCol ? progH.indexOf(groupCol) : -1;

            let groups = {};
            if (groupIdx >= 0) {
                progRows.forEach(r => {
                    const g = String(r[groupIdx] || 'Sin clasificar').trim();
                    if (!groups[g]) groups[g] = [];
                    groups[g].push(r);
                });
            } else {
                groups['Todos'] = progRows;
            }

            const content = [];
            const docTitle = state.pdfConfig.title || 'Informe de Programas Sociales';
            
            // Portada
            content.push({ 
                stack: [
                    { text: docTitle.toUpperCase(), style: 'docTitle' },
                    { text: 'CONSEJO NACIONAL DE COORDINACIÓN DE POLÍTICAS SOCIALES', style: 'docSub' }
                ],
                margin: [0, 150, 0, 0],
                pageBreak: 'after' 
            });

            const groupKeys = Object.keys(groups);
            const CONTENT_WIDTH = 515; // A4 pt aprox
            
            groupKeys.forEach((g, gIdx) => {
                const rows = groups[g];
                rows.forEach((progRow, pIdx) => {
                    // Start of program block
                    const progName = getVal(progRow, progH, 'Nombre del programa ingresado') ||
                        getVal(progRow, progH, 'Nombre del programa precargado') ||
                        `Programa ${pIdx + 1}`;

                    content.push({ text: progName.toUpperCase(), style: 'progTitle' });
                    content.push({
                        canvas: [{ type: 'line', x1: 0, y1: 0, x2: CONTENT_WIDTH, y2: 0, lineWidth: 2, lineColor: '#44658F' }],
                        margin: [0, 0, 0, 16]
                    });

                    const progFKVal = state.fkProgPrest ? getVal(progRow, progH, state.fkProgPrest) : null;
                    const progFKIndVal = state.fkProgInd ? getVal(progRow, progH, state.fkProgInd) : null;

                    state.sections.forEach(sec => {
                        if (sec.type === 'prog') {
                            content.push({ text: sec.title.toUpperCase(), style: 'secTitle' });
                            content.push({
                                canvas: [{ type: 'line', x1: 0, y1: 0, x2: CONTENT_WIDTH, y2: 0, lineWidth: 1, lineColor: '#cccccc' }],
                                margin: [0, 0, 0, 6]
                            });
                            sec.rows.forEach(r => {
                                if (r.sheet && r.sheet !== 'prog' && !r.type) return;
                                const val = getRowValue(r, progRow, null, null);
                                if (!val && r.type !== 'text-free') return;
                                injectFieldToPDF(r, val, content);
                            });
                        }
                        else if (sec.type === 'prest') {
                            const prestRows = state.data.prest || [];
                            const prestH = state.headers.prest || [];
                            const relPrest = state.fkProgPrest && progFKVal
                                ? prestRows.filter(pr => getVal(pr, prestH, state.fkProgPrest) === progFKVal)
                                : prestRows;
                                
                            if (relPrest.length === 0) return;
                            content.push({ text: sec.title.toUpperCase(), style: 'secTitle' });
                            content.push({
                                canvas: [{ type: 'line', x1: 0, y1: 0, x2: CONTENT_WIDTH, y2: 0, lineWidth: 1, lineColor: '#cccccc' }],
                                margin: [0, 0, 0, 6]
                            });

                            relPrest.forEach((pr, pi) => {
                                const prestName = getVal(pr, prestH, 'Prestación') || `Prestación ${pi + 1}`;
                                content.push({
                                    stack: [
                                        { text: `Prestación ${pi + 1}: ${prestName}`, style: 'prestSub' }
                                    ],
                                    margin: [0, 5, 0, 5]
                                });
                                
                                sec.rows.forEach(r => {
                                    if (r.sheet && r.sheet !== 'prest' && !r.type) return;
                                    const val = getRowValue(r, progRow, pr, null);
                                    if (!val && r.type !== 'text-free') return;
                                    injectFieldToPDF(r, val, content);
                                });

                                // Indicadores de esta prestación
                                const indSec = state.sections.find(s => s.type === 'ind');
                                if (indSec && indSec.rows.length > 0) {
                                    const indH = state.headers.ind || [];
                                    const indRows = state.data.ind || [];
                                    const prFKVal = state.fkPrestInd ? getVal(pr, prestH, state.fkPrestInd) : null;

                                    let relInds = indRows;
                                    if (prFKVal && state.fkPrestInd) {
                                        relInds = indRows.filter(ir => getVal(ir, indH, state.fkPrestInd) === prFKVal);
                                    } else if (progFKIndVal && state.fkProgInd) {
                                        relInds = indRows.filter(ir => getVal(ir, indH, state.fkProgInd) === progFKIndVal);
                                    }

                                    if (relInds.length > 0) {
                                        const tableCols = indSec.rows.filter(r => (r.sheet === 'ind' || r.type) && r.table);
                                        const hasTable = indSec.rows.some(r => r.table);

                                        if (hasTable && tableCols.length > 0) {
                                            const tHeader = tableCols.map(r => ({ text: r.label, style: 'th' }));
                                            const tBody = [tHeader];
                                            relInds.forEach(ir => {
                                                tBody.push(tableCols.map(r => ({ text: getRowValue(r, progRow, pr, ir) || 'S/I', style: 'td' })));
                                            });
                                            content.push({
                                                table: {
                                                    headerRows: 1,
                                                    widths: Array(tableCols.length).fill('*'),
                                                    body: tBody
                                                },
                                                layout: tableLayout,
                                                margin: [0, 6, 0, 12]
                                            });
                                        } else {
                                            content.push({ text: 'Indicadores:', style: 'tableTitle', margin: [0, 5, 0, 2] });
                                            relInds.forEach(ir => {
                                                indSec.rows.forEach(r => {
                                                    if (r.sheet && r.sheet !== 'ind' && !r.type) return;
                                                    const val = getRowValue(r, progRow, pr, ir);
                                                    if (!val && r.type !== 'text-free') return;
                                                    injectFieldToPDF(r, val, content);
                                                });
                                            });
                                        }
                                    }
                                }
                            });
                        }
                        else if (sec.type === 'ind') {
                            const hasPrestSec = state.sections.some(s => s.type === 'prest');
                            if (hasPrestSec) return;
                            
                            const indH = state.headers.ind || [];
                            const indRows = state.data.ind || [];
                            const relInds = state.fkProgInd && progFKIndVal
                                ? indRows.filter(ir => getVal(ir, indH, state.fkProgInd) === progFKIndVal)
                                : indRows;
                            if (relInds.length === 0) return;
                            
                            content.push({ text: sec.title.toUpperCase(), style: 'secTitle' });
                            content.push({
                                canvas: [{ type: 'line', x1: 0, y1: 0, x2: CONTENT_WIDTH, y2: 0, lineWidth: 1, lineColor: '#cccccc' }],
                                margin: [0, 0, 0, 6]
                            });

                            const tableCols = sec.rows.filter(r => (r.sheet === 'ind' || r.type) && r.table);
                            const hasTable = sec.rows.some(r => r.table);
                            
                            if (hasTable && tableCols.length > 0) {
                                const tHeader = tableCols.map(r => ({ text: r.label, style: 'th' }));
                                const tBody = [tHeader];
                                relInds.forEach(ir => {
                                    tBody.push(tableCols.map(r => ({ text: getRowValue(r, progRow, null, ir) || 'S/I', style: 'td' })));
                                });
                                content.push({
                                    table: {
                                        headerRows: 1,
                                        widths: Array(tableCols.length).fill('*'),
                                        body: tBody
                                    },
                                    layout: tableLayout,
                                    margin: [0, 6, 0, 12]
                                });
                            } else {
                                relInds.forEach(ir => {
                                    sec.rows.forEach(r => {
                                        if (r.sheet && r.sheet !== 'ind' && !r.type) return;
                                        const val = getRowValue(r, progRow, null, ir);
                                        if (!val && r.type !== 'text-free') return;
                                        injectFieldToPDF(r, val, content);
                                    });
                                });
                            }
                        }
                    });

                    // Page break after every program
                    if (pIdx < rows.length - 1 || gIdx < groupKeys.length - 1) {
                        content.push({ text: '', pageBreak: 'after' });
                    }
                });
            });

            return {
                pageMargins: [40, 105, 40, 40], // Margen superior a 105pt para compensar el nuevo alto de la imagen
                header: function(currentPage) {
                    // La portada (página 1) no lleva encabezado
                    if (currentPage === 1) return null;
                    if (headerImgBase64) {
                        return {
                            image: 'data:image/png;base64,' + headerImgBase64,
                            // Cálculo matemático exacto: 
                            // Ancho A4 (595.28pt) - Margen Izq (40pt) - Margen Der (40pt) = 515.28pt
                            width: 515.28, 
                            margin: [40, 35, 0, 0] // 35pt desde el borde superior, anclado al margen izquierdo
                        };
                    }
                    return null;
                },
                content: content,
                styles: {
                    docTitle: { fontSize: 30, bold: true, alignment: 'left', color: '#44658F', margin: [0, 0, 0, 10] },
                    docSub: { fontSize: 17, bold: true, alignment: 'left', color: '#5A5A5A', margin: [0, 0, 0, 10] },
                    groupTitle: { fontSize: 18, bold: true, color: '#44658F', margin: [0, 0, 0, 10] },
                    progTitle: { fontSize: 18, bold: true, color: '#44658F', alignment: 'left', margin: [0, 0, 0, 16] },
                    secTitle: { fontSize: 16, bold: true, color: '#44658F', margin: [0, 16, 0, 6] },
                    prestSub: { fontSize: 14, bold: true, color: '#000000', margin: [0, 5, 0, 5] },
                    tableTitle: { fontSize: 10, bold: true },
                    th: { bold: true, fontSize: 9, color: '#333333', fillColor: '#f0f0f0', alignment: 'left' },
                    td: { fontSize: 9, color: '#333333' },
                    fieldNormal: { fontSize: 10, margin: [0, 0, 0, 5], color: '#333333' },
                    fieldBold: { fontSize: 10, bold: true, margin: [0, 0, 0, 5], color: '#333333' },
                    fieldBig: { fontSize: 13, margin: [0, 0, 0, 5], color: '#333333' }
                },
                defaultStyle: {
                    font: 'Montserrat',
                    color: '#333333',
                    fontSize: 10
                }
            };
        }

        function injectFieldToPDF(row, val, contentArray) {
            let style = 'fieldNormal';
            if (row.bold) style = 'fieldBold';
            if (row.big) style = 'fieldBig';
            
            const labelStr = row.label ? `${row.label}: ` : '';

            if (row.bullet) {
                const items = String(val).split(/[,;]\s*/).filter(Boolean);
                contentArray.push({
                    stack: [
                        { text: row.label + ':', bold: true, fontSize: 10, margin: [0, 2, 0, 2] },
                        { ul: items.map(i => i.trim()), margin: [10, 0, 0, 5], fontSize: 10 }
                    ]
                });
            } else {
                contentArray.push({
                    text: [
                        { text: labelStr, bold: true },
                        String(val)
                    ],
                    style: style
                });
            }
        }

        // ══════════════════════════════════════════
        // UTILIDADES
        // ══════════════════════════════════════════
        function clearTemplate() {
            if (state.sections.length === 0) return;
            if (!confirm('¿Limpiar toda la plantilla?')) return;
            state.sections = [];
            sectionCounter = 0;
            renderTemplate();
            renderFieldList();
        }

        function updateStepPills(step) {
            ['pill-1', 'pill-2', 'pill-3'].forEach((id, i) => {
                const el = document.getElementById(id);
                el.className = 'step-pill';
                if (i + 1 < step) el.classList.add('done');
                else if (i + 1 === step) el.classList.add('active');
            });
        }

        function escAttr(s) {
            return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        // ══════════════════════════════════════════
        // PLANTILLA POR DEFECTO (si se carga el Excel conocido)
        // ══════════════════════════════════════════
        function buildDefaultTemplate() {
            // Solo si no hay secciones aún
            if (state.sections.length > 0) return;
            if (!state.headers.prog) return;

            const ph = state.headers.prog;
            const addSec = (title, type, cols) => {
                const id = 's' + (++sectionCounter);
                const rows = cols.flatMap(([sheet, col, label, fmt]) => {
                    const headers = state.headers[sheet] || [];
                    if (!headers.includes(col)) return [];
                    const r = { sheet, col, label: label || col, bold: false, big: false, bullet: false, table: false };
                    if (fmt) r[fmt] = true;
                    return [r];
                });
                state.sections.push({ id, title, type, rows });
            };

            addSec('Información Básica', 'prog', [
                ['prog', 'Jurisdicción', 'Jurisdicción'],
                ['prog', 'Ministerio/ Organismo descentralizado', 'Ministerio / Organismo'],
                ['prog', 'Secretaría', 'Secretaría'],
                ['prog', 'Subsecretaría', 'Subsecretaría'],
                ['prog', 'Autoridad del programa', 'Autoridad'],
                ['prog', 'Email del programa', 'Email'],
                ['prog', 'Página web', 'Página web'],
                ['prog', 'Domicilio programa', 'Domicilio'],
                ['prog', 'Teléfono de contacto', 'Teléfono'],
            ]);
            addSec('Características del Programa', 'prog', [
                ['prog', 'Temática', 'Temática'],
                ['prog', 'Sistema', 'Sistema'],
                ['prog', 'Objetivo general', 'Objetivo General', 'bold'],
                ['prog', 'Objetivo específico', 'Objetivos Específicos'],
                ['prog', 'Plan', 'Plan'],
                ['prog', 'Programa precedente', 'Programa Precedente'],
                ['prog', 'Normativa', 'Normativa'],
                ['prog', 'Alcance', 'Alcance'],
                ['prog', 'Población destinataria', 'Población Destinataria'],
                ['prog', 'Criterios de elegibilidad', 'Criterios de Elegibilidad'],
                ['prog', 'Modalidad de Ejecución', 'Modalidad'],
            ]);
            addSec('Prestaciones', 'prest', [
                ['prest', 'Tipo de prestación', 'Tipo'],
                ['prest', 'Requisitos de edad', 'Requisito de edad'],
            ]);
            addSec('Indicadores', 'ind', [
                ['ind', 'Indicador', 'Indicador', 'table'],
                ['ind', 'Unidad medida', 'Unidad', 'table'],
                ['ind', 'Periodicidad', 'Periodicidad', 'table'],
                ['ind', 'Ultimo dato disponible', 'Últ. Dato', 'table'],
                ['ind', 'Período', 'Período', 'table'],
                ['ind', 'Año', 'Año', 'table'],
            ]);

            renderTemplate();
            renderFieldList();
        }

        // Modificar loadExcel para auto-armar plantilla
        const origLoad = loadExcel;
        // Hook post-load
        const _origFileChange = fileInput.onchange;
        fileInput.addEventListener('change', () => {
            setTimeout(() => {
                if (state.headers.prog && state.sections.length === 0) buildDefaultTemplate();
            }, 500);
        });
        uploadZone.addEventListener('drop', () => {
            setTimeout(() => {
                if (state.headers.prog && state.sections.length === 0) buildDefaultTemplate();
            }, 500);
        });