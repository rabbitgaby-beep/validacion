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
            <button class="btn btn-secondary" style="padding: 4px 8px; font-size: 11px;" onclick="addSpecialRow('${sec.id}', 'text-free')">+ Texto Libre</button>
            <button class="btn btn-secondary" style="padding: 4px 8px; font-size: 11px;" onclick="addSpecialRow('${sec.id}', 'dropdown-map')">+ Campo Dinámico</button>
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

                const docDef = buildDocDefinition();
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

        function buildDocDefinition() {
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
            content.push({ text: docTitle, style: 'docTitle', pageBreak: 'after' });

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
                        canvas: [{ type: 'line', x1: 0, y1: 0, x2: CONTENT_WIDTH, y2: 0, lineWidth: 2, lineColor: '#003366' }],
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
                content: content,
                styles: {
                    docTitle: { fontSize: 24, bold: true, alignment: 'center', margin: [0, 200, 0, 0] },
                    groupTitle: { fontSize: 18, bold: true, color: '#003366', margin: [0, 0, 0, 10] },
                    progTitle: { fontSize: 16, bold: true, color: '#003366', alignment: 'center', margin: [0, 0, 0, 8] },
                    secTitle: { fontSize: 11, bold: true, color: '#004080', margin: [0, 16, 0, 3] },
                    prestSub: { fontSize: 11, bold: true, color: '#27ae60', margin: [0, 5, 0, 5] },
                    tableTitle: { fontSize: 10, bold: true },
                    th: { bold: true, fontSize: 9, color: '#333333', fillColor: '#f0f0f0', alignment: 'left' },
                    td: { fontSize: 9, color: '#333333' },
                    fieldNormal: { fontSize: 10, margin: [0, 0, 0, 5], color: '#222222' },
                    fieldBold: { fontSize: 10, bold: true, margin: [0, 0, 0, 5], color: '#222222' },
                    fieldBig: { fontSize: 13, margin: [0, 0, 0, 5], color: '#222222' }
                },
                defaultStyle: {
                    font: 'Montserrat',
                    color: '#222222'
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