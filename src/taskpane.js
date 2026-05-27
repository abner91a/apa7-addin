/* ================================================================
   Referencias APA 7 — Lógica del complemento
   ================================================================ */

// ---- Office init ----
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initApp();
    }
});

// ---- Estado global ----
const STORAGE_KEY = 'apa7_refs_v1';
let savedRefs = loadRefs();

// ---- Inicialización ----
function initApp() {
    setupTabs();
    setupTypeChange();
    setupActions();
    renderFields(); // pinta los campos iniciales
    renderList();
    updateBadge();
}

// =================================================================
// TABS
// =================================================================
function setupTabs() {
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
            tab.classList.add('active');
            document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
        });
    });
}

// =================================================================
// CAMPOS DINÁMICOS por tipo de fuente
// =================================================================
const FIELDS_BY_TYPE = {
    libro: [
        { id: 'year', label: 'Año', placeholder: '2023', type: 'text', size: 'sm' },
        { id: 'title', label: 'Título del libro', placeholder: 'Cien años de soledad', type: 'text' },
        { id: 'edition', label: 'Edición (opcional)', placeholder: '3ª ed.', type: 'text', size: 'sm' },
        { id: 'publisher', label: 'Editorial', placeholder: 'Editorial Sudamericana', type: 'text' },
        { id: 'doi', label: 'DOI o URL (opcional)', placeholder: 'https://doi.org/10.xxxx/xxxx', type: 'text' }
    ],
    capitulo: [
        { id: 'year', label: 'Año', placeholder: '2023', type: 'text', size: 'sm' },
        { id: 'chapTitle', label: 'Título del capítulo', placeholder: 'El método científico', type: 'text' },
        { id: 'editors', label: 'Editores del libro', placeholder: 'Pérez, Juan; López, Ana', type: 'text', hint: 'Uno por línea o separados por punto y coma. Formato: Apellido, Nombre' },
        { id: 'bookTitle', label: 'Título del libro', placeholder: 'Manual de metodología', type: 'text' },
        { id: 'pages', label: 'Páginas (pp.)', placeholder: '45-72', type: 'text', size: 'sm' },
        { id: 'edition', label: 'Edición (opcional)', placeholder: '2ª ed.', type: 'text', size: 'sm' },
        { id: 'publisher', label: 'Editorial', placeholder: 'Siglo XXI', type: 'text' },
        { id: 'doi', label: 'DOI o URL (opcional)', placeholder: 'https://doi.org/...', type: 'text' }
    ],
    articulo: [
        { id: 'year', label: 'Año', placeholder: '2023', type: 'text', size: 'sm' },
        { id: 'title', label: 'Título del artículo', placeholder: 'Impacto del cambio climático en...', type: 'text' },
        { id: 'journal', label: 'Nombre de la revista', placeholder: 'Revista de Ciencias Sociales', type: 'text' },
        { id: 'volume', label: 'Volumen', placeholder: '24', type: 'text', size: 'sm' },
        { id: 'issue', label: 'Número (opcional)', placeholder: '3', type: 'text', size: 'sm' },
        { id: 'pages', label: 'Páginas', placeholder: '125-148', type: 'text', size: 'sm' },
        { id: 'doi', label: 'DOI o URL (recomendado)', placeholder: 'https://doi.org/10.xxxx/xxxx', type: 'text' }
    ],
    periodico: [
        { id: 'date', label: 'Fecha completa', placeholder: '2023, 15 de marzo', type: 'text', hint: 'Formato: AAAA, día de mes' },
        { id: 'title', label: 'Título del artículo', placeholder: 'Avances en energía renovable', type: 'text' },
        { id: 'newspaper', label: 'Nombre del periódico', placeholder: 'El País', type: 'text' },
        { id: 'pages', label: 'Páginas (opcional)', placeholder: 'A1, A4', type: 'text', size: 'sm' },
        { id: 'url', label: 'URL (si es en línea)', placeholder: 'https://elpais.com/...', type: 'text' }
    ],
    web: [
        { id: 'date', label: 'Fecha de publicación', placeholder: '2023, 12 de junio', type: 'text', hint: 'AAAA, día de mes. Usa "s.f." si no hay fecha' },
        { id: 'title', label: 'Título de la página', placeholder: 'Guía sobre vacunación', type: 'text' },
        { id: 'site', label: 'Nombre del sitio web', placeholder: 'Organización Mundial de la Salud', type: 'text', hint: 'Omitir si es igual al autor' },
        { id: 'url', label: 'URL', placeholder: 'https://www.who.int/...', type: 'text' }
    ],
    tesis: [
        { id: 'year', label: 'Año', placeholder: '2023', type: 'text', size: 'sm' },
        { id: 'title', label: 'Título de la tesis', placeholder: 'Estudio sobre...', type: 'text' },
        { id: 'thesisType', label: 'Tipo', placeholder: 'Tesis doctoral', type: 'text', hint: 'Ej: Tesis de licenciatura, Tesis de maestría, Tesis doctoral' },
        { id: 'institution', label: 'Institución', placeholder: 'Universidad Nacional Autónoma de México', type: 'text' },
        { id: 'repository', label: 'Repositorio (opcional)', placeholder: 'Repositorio Institucional UNAM', type: 'text' },
        { id: 'url', label: 'URL (opcional)', placeholder: 'https://...', type: 'text' }
    ],
    informe: [
        { id: 'year', label: 'Año', placeholder: '2023', type: 'text', size: 'sm' },
        { id: 'title', label: 'Título del informe', placeholder: 'Informe sobre desarrollo humano', type: 'text' },
        { id: 'number', label: 'Número de informe (opcional)', placeholder: 'No. 45', type: 'text', size: 'sm' },
        { id: 'publisher', label: 'Editorial / Organismo', placeholder: 'PNUD', type: 'text', hint: 'Omitir si es igual al autor' },
        { id: 'url', label: 'URL (opcional)', placeholder: 'https://...', type: 'text' }
    ],
    video: [
        { id: 'date', label: 'Fecha', placeholder: '2023, 5 de mayo', type: 'text', hint: 'AAAA, día de mes' },
        { id: 'title', label: 'Título del video', placeholder: 'Cómo funciona la fotosíntesis', type: 'text' },
        { id: 'platform', label: 'Plataforma', placeholder: 'YouTube', type: 'text', size: 'sm' },
        { id: 'url', label: 'URL', placeholder: 'https://youtube.com/watch?v=...', type: 'text' }
    ]
};

function setupTypeChange() {
    document.getElementById('ref-type').addEventListener('change', renderFields);
}

function renderFields() {
    const type = document.getElementById('ref-type').value;
    const fields = FIELDS_BY_TYPE[type];
    const wrap = document.getElementById('dynamic-fields');
    wrap.innerHTML = '';

    fields.forEach(f => {
        const block = document.createElement('div');
        block.className = 'field-block';

        const label = document.createElement('label');
        label.className = 'field-label';
        label.textContent = f.label;
        label.setAttribute('for', 'f-' + f.id);
        block.appendChild(label);

        if (f.hint) {
            const hint = document.createElement('p');
            hint.className = 'hint';
            hint.innerHTML = f.hint;
            block.appendChild(hint);
        }

        const input = document.createElement('input');
        input.type = f.type;
        input.id = 'f-' + f.id;
        input.className = 'input';
        input.placeholder = f.placeholder || '';
        input.autocomplete = 'off';
        block.appendChild(input);

        wrap.appendChild(block);
    });

    // limpiar vista previa al cambiar tipo
    document.getElementById('preview-wrap').hidden = true;
    setStatus('');
}

// =================================================================
// PARSEO DE AUTORES
// =================================================================
/**
 * Parsea el campo de autores. Cada línea (o separado por ;) es un autor.
 * Formatos aceptados:
 *   "Apellido, Nombre Segundo"   → { surname: "Apellido", initials: "N. S." }
 *   "[Institución]"               → { group: "Institución" }
 *   "[Organización Mundial de la Salud] OMS"  → { group: "..." }
 */
function parseAuthors(raw) {
    if (!raw || !raw.trim()) return [];
    const lines = raw.split(/\n|;/).map(s => s.trim()).filter(Boolean);

    return lines.map(line => {
        // Autor institucional
        if (line.startsWith('[')) {
            const m = line.match(/^\[(.+?)\]/);
            return { group: m ? m[1].trim() : line.replace(/[\[\]]/g, '').trim() };
        }
        // Autor persona: "Apellido(s), Nombre(s)"
        const parts = line.split(',');
        if (parts.length < 2) {
            // sin coma: tratamos todo como apellido único
            return { surname: line.trim(), initials: '' };
        }
        const surname = parts[0].trim();
        const givenNames = parts.slice(1).join(',').trim();
        const initials = givenNames
            .split(/\s+/)
            .filter(Boolean)
            .map(n => n.charAt(0).toUpperCase() + '.')
            .join(' ');
        return { surname, initials };
    });
}

/**
 * Formatea autores para la lista de referencias APA 7.
 * Devuelve un string (sin cursivas).
 */
function formatAuthorsForReference(authors) {
    if (authors.length === 0) return '';

    const fmt = (a) => a.group ? a.group : `${a.surname}, ${a.initials}`.trim();

    if (authors.length === 1) return fmt(authors[0]);

    if (authors.length <= 20) {
        const all = authors.map(fmt);
        const last = all.pop();
        return all.join(', ') + ', y ' + last;
    }

    // 21+ autores: primeros 19, elipsis, último
    const first19 = authors.slice(0, 19).map(fmt);
    const lastAuthor = fmt(authors[authors.length - 1]);
    return first19.join(', ') + ', ... ' + lastAuthor;
}

/**
 * Formatea autores para cita en texto.
 */
function formatAuthorsInText(authors, year, narrative = false) {
    if (authors.length === 0) return '';

    const surn = (a) => a.group || a.surname;

    let names;
    if (authors.length === 1) {
        names = surn(authors[0]);
    } else if (authors.length === 2) {
        names = narrative
            ? `${surn(authors[0])} y ${surn(authors[1])}`
            : `${surn(authors[0])} y ${surn(authors[1])}`;
    } else {
        names = `${surn(authors[0])} et al.`;
    }

    return narrative ? `${names} (${year})` : `(${names}, ${year})`;
}

// =================================================================
// FORMATEADORES POR TIPO → devuelven [{text, italic}]
// =================================================================
function buildReference(type, authors, data) {
    switch (type) {
        case 'libro':       return buildLibro(authors, data);
        case 'capitulo':    return buildCapitulo(authors, data);
        case 'articulo':    return buildArticulo(authors, data);
        case 'periodico':   return buildPeriodico(authors, data);
        case 'web':         return buildWeb(authors, data);
        case 'tesis':       return buildTesis(authors, data);
        case 'informe':     return buildInforme(authors, data);
        case 'video':       return buildVideo(authors, data);
        default: return [];
    }
}

// Helper para construir segmentos
const seg = (text, italic = false) => ({ text, italic });

function buildLibro(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.year || 's.f.'}). `));
    out.push(seg(d.title || '[Título desconocido]', true));
    if (d.edition) out.push(seg(` (${d.edition})`));
    out.push(seg('. '));
    if (d.publisher) out.push(seg(`${d.publisher}.`));
    if (d.doi) out.push(seg(` ${d.doi}`));
    return out;
}

function buildCapitulo(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.year || 's.f.'}). `));
    out.push(seg(`${d.chapTitle || '[Título del capítulo]'}. `));

    // Editores
    let editorsStr = '';
    if (d.editors) {
        const eds = parseAuthors(d.editors);
        const fmt = (a) => a.group ? a.group : `${a.initials} ${a.surname}`.trim();
        if (eds.length === 1) editorsStr = `En ${fmt(eds[0])} (Ed.), `;
        else if (eds.length > 1) {
            const all = eds.map(fmt);
            const last = all.pop();
            editorsStr = `En ${all.join(', ')} y ${last} (Eds.), `;
        }
    }
    out.push(seg(editorsStr));
    out.push(seg(d.bookTitle || '[Título del libro]', true));

    // Edición + páginas en paréntesis
    const pars = [];
    if (d.edition) pars.push(d.edition);
    if (d.pages) pars.push(`pp. ${d.pages}`);
    if (pars.length) out.push(seg(` (${pars.join(', ')})`));
    out.push(seg('. '));

    if (d.publisher) out.push(seg(`${d.publisher}.`));
    if (d.doi) out.push(seg(` ${d.doi}`));
    return out;
}

function buildArticulo(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.year || 's.f.'}). `));
    out.push(seg(`${d.title || '[Título del artículo]'}. `));
    out.push(seg(d.journal || '[Revista]', true));
    if (d.volume) {
        out.push(seg(', '));
        out.push(seg(d.volume, true));
        if (d.issue) out.push(seg(`(${d.issue})`));
    }
    if (d.pages) out.push(seg(`, ${d.pages}`));
    out.push(seg('.'));
    if (d.doi) out.push(seg(` ${d.doi}`));
    return out;
}

function buildPeriodico(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.date || 's.f.'}). `));
    out.push(seg(`${d.title || '[Título del artículo]'}. `));
    out.push(seg(d.newspaper || '[Periódico]', true));
    if (d.pages) out.push(seg(`, ${d.pages}`));
    out.push(seg('.'));
    if (d.url) out.push(seg(` ${d.url}`));
    return out;
}

function buildWeb(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.date || 's.f.'}). `));
    out.push(seg(d.title || '[Título]', true));
    out.push(seg('. '));
    // Si el sitio es diferente del autor, lo incluimos
    if (d.site) {
        const firstAuthor = authors[0];
        const authorName = firstAuthor ? (firstAuthor.group || firstAuthor.surname) : '';
        if (d.site.trim().toLowerCase() !== authorName.trim().toLowerCase()) {
            out.push(seg(`${d.site}. `));
        }
    }
    if (d.url) out.push(seg(d.url));
    return out;
}

function buildTesis(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.year || 's.f.'}). `));
    out.push(seg(d.title || '[Título de la tesis]', true));
    const bracket = [];
    if (d.thesisType) bracket.push(d.thesisType);
    if (d.institution) bracket.push(d.institution);
    if (bracket.length) out.push(seg(` [${bracket.join(', ')}]`));
    out.push(seg('. '));
    if (d.repository) out.push(seg(`${d.repository}. `));
    if (d.url) out.push(seg(d.url));
    return out;
}

function buildInforme(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.year || 's.f.'}). `));
    out.push(seg(d.title || '[Título del informe]', true));
    if (d.number) out.push(seg(` (${d.number})`));
    out.push(seg('. '));
    if (d.publisher) {
        const firstAuthor = authors[0];
        const authorName = firstAuthor ? (firstAuthor.group || firstAuthor.surname) : '';
        if (d.publisher.trim().toLowerCase() !== authorName.trim().toLowerCase()) {
            out.push(seg(`${d.publisher}. `));
        }
    }
    if (d.url) out.push(seg(d.url));
    return out;
}

function buildVideo(authors, d) {
    const out = [];
    const authStr = formatAuthorsForReference(authors);
    out.push(seg(`${authStr} (${d.date || 's.f.'}). `));
    out.push(seg(d.title || '[Título del video]', true));
    out.push(seg(' [Video]. '));
    if (d.platform) out.push(seg(`${d.platform}. `));
    if (d.url) out.push(seg(d.url));
    return out;
}

// =================================================================
// HELPERS de UI
// =================================================================
function getFormData() {
    const type = document.getElementById('ref-type').value;
    const authorsRaw = document.getElementById('authors').value;
    const authors = parseAuthors(authorsRaw);

    const data = {};
    FIELDS_BY_TYPE[type].forEach(f => {
        const el = document.getElementById('f-' + f.id);
        if (el) data[f.id] = el.value.trim();
    });

    return { type, authors, data, authorsRaw };
}

function segmentsToPlain(segs) {
    return segs.map(s => s.text).join('');
}

function segmentsToHTML(segs) {
    return segs.map(s => {
        const esc = s.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        return s.italic ? `<em>${esc}</em>` : esc;
    }).join('');
}

function getYearForCitation(data) {
    return data.year || (data.date ? (data.date.split(',')[0] || 's.f.') : 's.f.');
}

// =================================================================
// ACCIONES (botones)
// =================================================================
function setupActions() {
    document.getElementById('btn-preview').addEventListener('click', onPreview);
    document.getElementById('btn-insert-ref').addEventListener('click', onInsertRef);
    document.getElementById('btn-insert-cite').addEventListener('click', onInsertCite);
    document.getElementById('btn-save').addEventListener('click', onSave);
    document.getElementById('btn-clear-all').addEventListener('click', onClearAll);
    document.getElementById('btn-insert-bibliography').addEventListener('click', onInsertBibliography);
}

function onPreview() {
    const { type, authors, data } = getFormData();
    if (authors.length === 0) {
        setStatus('Agrega al menos un autor.', true);
        return;
    }
    const segs = buildReference(type, authors, data);
    const year = getYearForCitation(data);

    document.getElementById('preview-ref').innerHTML = segmentsToHTML(segs);
    document.getElementById('preview-intext').textContent = formatAuthorsInText(authors, year, false);
    document.getElementById('preview-intext-narr').textContent = 'Narrativa: ' + formatAuthorsInText(authors, year, true);
    document.getElementById('preview-wrap').hidden = false;
    setStatus('Vista previa actualizada.');
}

async function onInsertRef() {
    const { type, authors, data } = getFormData();
    if (authors.length === 0) { setStatus('Agrega al menos un autor.', true); return; }
    const segs = buildReference(type, authors, data);

    try {
        await insertSegmentsAsParagraph(segs);
        setStatus('Referencia insertada.');
    } catch (err) {
        setStatus('Error al insertar: ' + err.message, true);
    }
}

async function onInsertCite() {
    const { authors, data } = getFormData();
    if (authors.length === 0) { setStatus('Agrega al menos un autor.', true); return; }
    const year = getYearForCitation(data);
    const cite = formatAuthorsInText(authors, year, false);

    try {
        await insertText(cite);
        setStatus('Cita insertada en el cursor.');
    } catch (err) {
        setStatus('Error al insertar: ' + err.message, true);
    }
}

function onSave() {
    const { type, authors, data, authorsRaw } = getFormData();
    if (authors.length === 0) { setStatus('Agrega al menos un autor antes de guardar.', true); return; }
    const segs = buildReference(type, authors, data);
    const sortKey = (authors[0].group || authors[0].surname || 'zzz').toLowerCase();

    savedRefs.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        type, authorsRaw, data, segments: segs, sortKey
    });
    saveRefs();
    renderList();
    updateBadge();
    setStatus(`Guardada en lista (${savedRefs.length} en total).`);
}

function onClearAll() {
    if (savedRefs.length === 0) return;
    if (!confirm('¿Vaciar la lista completa de referencias guardadas?')) return;
    savedRefs = [];
    saveRefs();
    renderList();
    updateBadge();
}

async function onInsertBibliography() {
    if (savedRefs.length === 0) { setStatus('La lista está vacía.', true); return; }

    // Ordenar alfabéticamente por sortKey
    const sorted = [...savedRefs].sort((a, b) =>
        a.sortKey.localeCompare(b.sortKey, 'es', { sensitivity: 'base' })
    );

    try {
        await insertBibliography(sorted);
        setStatus(`Insertadas ${sorted.length} referencias al final del documento.`);
    } catch (err) {
        setStatus('Error: ' + err.message, true);
    }
}

// =================================================================
// LISTA UI
// =================================================================
function renderList() {
    const wrap = document.getElementById('ref-list');
    const countEl = document.getElementById('list-count');
    wrap.innerHTML = '';

    if (savedRefs.length === 0) {
        wrap.innerHTML = '<div class="empty-state">Aún no guardas referencias.<br>Créalas en la pestaña <strong>Crear</strong>.</div>';
        countEl.textContent = '0 referencias guardadas';
        return;
    }

    countEl.textContent = `${savedRefs.length} ${savedRefs.length === 1 ? 'referencia guardada' : 'referencias guardadas'}`;

    savedRefs.forEach(r => {
        const item = document.createElement('div');
        item.className = 'ref-item';
        item.innerHTML = `
            <div class="ref-item-text">${segmentsToHTML(r.segments)}</div>
            <div class="ref-item-actions">
                <button data-action="insert" data-id="${r.id}">Insertar</button>
                <button data-action="delete" data-id="${r.id}">Eliminar</button>
            </div>
        `;
        wrap.appendChild(item);
    });

    wrap.querySelectorAll('button[data-action]').forEach(b => {
        b.addEventListener('click', async () => {
            const id = b.dataset.id;
            const ref = savedRefs.find(r => r.id === id);
            if (!ref) return;
            if (b.dataset.action === 'insert') {
                try {
                    await insertSegmentsAsParagraph(ref.segments);
                    setStatus('Referencia insertada.');
                } catch (e) { setStatus(e.message, true); }
            } else {
                savedRefs = savedRefs.filter(r => r.id !== id);
                saveRefs();
                renderList();
                updateBadge();
            }
        });
    });
}

function updateBadge() {
    document.getElementById('badge-count').textContent = savedRefs.length;
}

// =================================================================
// PERSISTENCIA (localStorage)
// =================================================================
function loadRefs() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        return raw ? JSON.parse(raw) : [];
    } catch { return []; }
}

function saveRefs() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(savedRefs)); }
    catch (e) { console.warn('No se pudo guardar en localStorage', e); }
}

// =================================================================
// STATUS
// =================================================================
function setStatus(msg, isError = false) {
    const el = document.getElementById('status');
    el.textContent = msg;
    el.classList.toggle('error', isError);
    if (msg && !isError) {
        clearTimeout(setStatus._t);
        setStatus._t = setTimeout(() => { el.textContent = ''; }, 4000);
    }
}

// =================================================================
// OFFICE.JS — Inserción en Word
// =================================================================

/**
 * Inserta texto plano en la posición del cursor.
 */
async function insertText(text) {
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(text, Word.InsertLocation.replace);
        await context.sync();
    });
}

/**
 * Inserta los segmentos como un párrafo nuevo con sangría francesa (hanging),
 * aplicando cursiva donde corresponda.
 */
async function insertSegmentsAsParagraph(segments) {
    await Word.run(async (context) => {
        const sel = context.document.getSelection();
        // Insertar un nuevo párrafo vacío después de la posición actual
        const paragraph = sel.insertParagraph('', Word.InsertLocation.after);
        // Sangría francesa: indentamos el párrafo y luego el primer renglón se "saca"
        paragraph.leftIndent = 36;          // 0.5"  ≈ 36 pt
        paragraph.firstLineIndent = -36;     // sangría francesa
        paragraph.spaceAfter = 0;

        // Insertar cada segmento como un range con su formato
        let currentRange = paragraph.getRange(Word.RangeLocation.end);
        for (const s of segments) {
            const inserted = currentRange.insertText(s.text, Word.InsertLocation.after);
            if (s.italic) inserted.font.italic = true;
            currentRange = inserted.getRange(Word.RangeLocation.end);
        }

        await context.sync();
    });
}

/**
 * Inserta toda la bibliografía al final del documento con título "Referencias".
 */
async function insertBibliography(refs) {
    await Word.run(async (context) => {
        const body = context.document.body;

        // Salto y título
        const titlePara = body.insertParagraph('Referencias', Word.InsertLocation.end);
        titlePara.alignment = Word.Alignment.centered;
        titlePara.font.bold = true;
        titlePara.font.size = 12;
        titlePara.spaceAfter = 12;

        // Insertar cada referencia
        for (const r of refs) {
            const p = body.insertParagraph('', Word.InsertLocation.end);
            p.leftIndent = 36;
            p.firstLineIndent = -36;
            p.spaceAfter = 0;

            let rng = p.getRange(Word.RangeLocation.end);
            for (const s of r.segments) {
                const ins = rng.insertText(s.text, Word.InsertLocation.after);
                if (s.italic) ins.font.italic = true;
                rng = ins.getRange(Word.RangeLocation.end);
            }
        }
        await context.sync();
    });
}
