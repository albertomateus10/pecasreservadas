/* ===== Registro do Plugin de Rótulos do Chart.js ===== */
if (typeof Chart !== 'undefined' && typeof ChartDataLabels !== 'undefined') {
    Chart.register(ChartDataLabels);
}

/* ===== Helpers UI ===== */
function showError(msg) {
    const el = document.getElementById('alertError');
    el.querySelector('.msg').textContent = msg;
    el.classList.add('show');
}
function showWarn(msg) {
    const el = document.getElementById('alertWarn');
    el.querySelector('.msg').textContent = msg;
    el.classList.add('show');
}

/* ===== Util ===== */
const $ = s => document.querySelector(s);
const NUM = v => new Intl.NumberFormat('pt-BR').format(+v || 0);
const DEC = v => new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(+v || 0);
const PCT = v => (Number.isFinite(v) ? v.toFixed(1).replace('.', ',') : '—') + '%';
const BRL = v => (+v ? new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(+v) : '—');
const escAttr = s => String(s ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/* ===== Scroll estável ===== */
function withStableScroll(fn) {
    const x = window.pageXOffset || 0;
    const y = window.pageYOffset || 0;
    fn();
    requestAnimationFrame(() => {
        window.scrollTo(x, y);
        requestAnimationFrame(() => { window.scrollTo(x, y); });
    });
}

/* ===== Conversões & Datas ===== */
function toNumberBR(x) {
    if (x === null || x === undefined || x === '') return 0;
    if (typeof x === 'number') return x;
    const s = String(x).trim();
    const clean = s.replace(/[R$\s]/g, '').replace(/\./g, '').replace(/,/g, '.');
    const n = parseFloat(clean);
    return isNaN(n) ? 0 : n;
}

function parseDateExcel(v) {
    if (typeof v === 'number') {
        const base = new Date(Date.UTC(1899, 11, 30));
        return new Date(base.getTime() + v * 86400000);
    }
    if (typeof v === 'string') {
        // Format: DD/MM/YYYY or DD/MM/YYYY HH:MM:SS
        const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (m) {
            return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
        }
    }
    const d = new Date(v);
    return isNaN(+d) ? null : d;
}

const MONTH_NAMES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

function monthYearFromCell(cell) {
    if (cell === null || cell === undefined || cell === '' || cell === '—') return 'Não informado';

    if (typeof cell === 'string') {
        const m = cell.match(/\/(\d{2})\/(\d{4})/);
        if (m) {
            const mm = parseInt(m[1], 10);
            const aaaa = m[2];
            if (mm >= 1 && mm <= 12) {
                return `${MONTH_NAMES[mm - 1]} ${aaaa}`;
            }
        }
    }

    const d = parseDateExcel(cell);
    if (d && !isNaN(+d)) {
        return `${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}`;
    }
    return 'Não informado';
}

function mapFaixaDias(dias) {
    const d = Number(dias) || 0;
    if (d <= 30) return 'de 0 a 30 dias';
    if (d <= 90) return 'de 31 a 90 dias';
    if (d <= 180) return 'de 91 a 180 dias';
    if (d <= 365) return 'de 181 a 365 dias';
    return 'acima de 365 dias';
}

function normalizeOrigem(raw) {
    const s = String(raw || '').trim();
    if (!s) return 'Não informado';
    if (s.includes('Ordem') || s.includes('Servico') || s.includes('Serviço')) return 'Ordem de Serviço';
    if (s.includes('Autom')) return 'Automática';
    if (s.includes('Orcamento') || s.includes('Orçamento')) {
        if (s.includes('Interno')) return 'Orçamento Interno';
        if (s.includes('Externo')) return 'Orçamento Externo';
        return 'Orçamento';
    }
    if (s.includes('Pr') && s.includes('Nota')) return 'Pré-Nota Completa';
    return s;
}

function normalizeLiberacao(raw) {
    const s = String(raw || '').trim().toUpperCase();
    if (s === 'S' || s === 'SIM' || s === 'LIBERADO') return 'Liberada';
    if (s === 'N' || s === 'NÃO' || s === 'NAO' || s === 'BLOQUEADO') return 'Não Liberada';
    return 'Pendente / Não Informado';
}

/* ===== Estado da Aplicação ===== */
let allData = [];
let filtered = [];
let charts = {
    periodo: null,
    topPecas: null,
    vendedor: null,
    origem: null,
    status: null,
    dias: null
};

const state = {
    search: '',
    Mes: [],
    Origem: [],
    Vendedor: [],
    Codigo: [],
    Prioridade: [],
    Liberacao: [],
    StatusOS: [],
    FaixaDias: []
};

const sortState = { key: null, type: null, dir: 'desc' };

/* Estado da Tabela sob Demanda */
let tableLoaded = false;
let barFilter = null; // { type: string, label: string }
let currentPage = 1;
const pageSize = 50;

/* ===== Componente MultiPill ===== */
class MultiPill {
    constructor({ label, getter, setter }) {
        this.label = label;
        this.getter = getter;
        this.setter = setter;
        this.options = [];

        this.el = document.createElement('div');
        this.el.className = 'pill';
        this.el.tabIndex = 0;
        this.el.innerHTML = `<strong>${label}:</strong> <span class="value">Todos</span> <i class="fa-solid fa-chevron-down chev"></i><div class="menu"></div>`;
        this.menu = this.el.querySelector('.menu');
        this.valueEl = this.el.querySelector('.value');
        this.searchInput = null;

        this.el.addEventListener('click', e => {
            if (!e.target.closest('.menu')) {
                const x = window.pageXOffset, y = window.pageYOffset;
                const wasOpen = this.el.classList.contains('open');
                document.querySelectorAll('.pill.open').forEach(p => p.classList.remove('open'));
                if (!wasOpen) {
                    this.el.classList.add('open');
                    if (this.searchInput) {
                        setTimeout(() => this.searchInput.focus(), 50);
                    }
                }
                window.scrollTo(x, y);
            }
        });

        document.addEventListener('click', e => {
            if (!this.el.contains(e.target)) {
                this.el.classList.remove('open');
            }
        });
    }

    setOptions(opts, { keepOrder = false } = {}) {
        const unique = [...new Set(opts.filter(Boolean))];
        this.options = keepOrder ? unique : unique.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { numeric: true }));
        this.render();
    }

    render() {
        const sel = new Set(this.getter());
        this.menu.innerHTML = '';

        const s = document.createElement('input');
        s.placeholder = `Pesquisar ${this.label.toLowerCase()}...`;
        s.className = 'menu-search';
        this.menu.append(s);
        this.searchInput = s;

        const acts = document.createElement('div');
        acts.className = 'menu-actions';
        const limpar = document.createElement('button');
        limpar.textContent = 'Limpar';
        const aplicar = document.createElement('button');
        aplicar.textContent = 'Aplicar';
        aplicar.className = 'apply';
        acts.append(limpar, aplicar);
        this.menu.append(acts);

        const list = document.createElement('div');
        list.className = 'list';
        this.menu.append(list);

        for (const o of this.options) {
            const row = document.createElement('label');
            row.className = 'opt';
            row.dataset.text = String(o).toLowerCase();
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.value = o;
            chk.checked = sel.has(String(o));
            row.append(chk, document.createElement('span'));
            row.lastChild.textContent = o;
            list.append(row);
        }

        s.oninput = () => {
            const q = s.value.toLowerCase();
            list.querySelectorAll('.opt').forEach(el => {
                el.style.display = el.dataset.text.includes(q) ? '' : 'none';
            });
        };

        limpar.onclick = (e) => {
            e.stopPropagation();
            withStableScroll(() => {
                this.setter([]);
                list.querySelectorAll('input:checked').forEach(c => c.checked = false);
                this.sync();
                applyFilters();
            });
        };

        aplicar.onclick = (e) => {
            e.stopPropagation();
            withStableScroll(() => {
                const vals = [...list.querySelectorAll('input:checked')].map(i => i.value);
                this.setter(vals);
                this.el.classList.remove('open');
                this.sync();
                applyFilters();
            });
        };

        list.addEventListener('click', (ev) => {
            const input = ev.target.closest('input[type="checkbox"]');
            if (!input) return;
            const temp = new Set(this.getter());
            if (input.checked) temp.add(input.value); else temp.delete(input.value);
            const arr = [...temp];
            this.valueEl.textContent = !arr.length ? 'Todos' : (arr.length <= 2 ? arr.join(', ') : `${arr[0]}, ${arr[1]} (+${arr.length - 2})`);
        });

        this.sync();
    }

    sync() {
        const v = this.getter();
        this.valueEl.textContent = !v || v.length === 0 ? 'Todos' : (v.length <= 2 ? v.join(', ') : `${v[0]}, ${v[1]} (+${v.length - 2})`);
    }
}

/* Instanciação das Pílulas */
const pillInstances = {
    Mes: new MultiPill({ label: 'Período', getter: () => state.Mes, setter: v => state.Mes = v }),
    Origem: new MultiPill({ label: 'Origem', getter: () => state.Origem, setter: v => state.Origem = v }),
    Vendedor: new MultiPill({ label: 'Vendedor', getter: () => state.Vendedor, setter: v => state.Vendedor = v }),
    Codigo: new MultiPill({ label: 'Código da Peça', getter: () => state.Codigo, setter: v => state.Codigo = v }),
    Prioridade: new MultiPill({ label: 'Prioridade', getter: () => state.Prioridade, setter: v => state.Prioridade = v }),
    Liberacao: new MultiPill({ label: 'Liberação', getter: () => state.Liberacao, setter: v => state.Liberacao = v }),
    StatusOS: new MultiPill({ label: 'Status OS/Orç', getter: () => state.StatusOS, setter: v => state.StatusOS = v }),
    FaixaDias: new MultiPill({ label: 'Faixa de Dias', getter: () => state.FaixaDias, setter: v => state.FaixaDias = v }),
};

(function mountPills() {
    const holder = document.getElementById('filters');
    if (!holder) return;
    ['Mes', 'Origem', 'Vendedor', 'Codigo', 'Prioridade', 'Liberacao', 'StatusOS', 'FaixaDias'].forEach(k => {
        holder.appendChild(pillInstances[k].el);
    });
})();

/* ===== Upload de Planilha ===== */
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
if (fileInput) {
    fileInput.addEventListener('change', ev => {
        const f = ev.target.files?.[0];
        if (f) {
            fileNameEl.textContent = f.name;
            readFile(f);
        }
    });
}

async function readFile(file) {
    try {
        if (!window.XLSX) {
            showError('Biblioteca XLSX não carregada. Verifique sua conexão com a internet.');
            return;
        }
        const ext = (file.name.split('.').pop() || '').toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(ext)) {
            showWarn('Formato não suportado. Por favor utilize arquivos .xlsx, .xls ou .csv.');
            return;
        }
        if (ext === 'csv') {
            const text = await file.text();
            const wb = XLSX.read(text, { type: 'string' });
            processWorkbook(wb);
        } else {
            const buf = await file.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
            processWorkbook(wb);
        }
    } catch (err) {
        console.error(err);
        showError('Falha ao ler o arquivo selecionado. Verifique se a planilha possui dados válidos.');
    }
}

/* =====================================================================
 * LEITURA DA PLANILHA — MAPEAMENTO DE COLUNAS PELO CABEÇALHO
 * ---------------------------------------------------------------------
 * A extração da base tem o seguinte layout:
 *   Linha 1 → título do relatório
 *   Linha 2 → "Data de Geração: ..."
 *   Linha 3 → CABEÇALHO (nomes das colunas)
 *   Linha 4 em diante → dados
 *
 * Cada campo é localizado pelo NOME do cabeçalho da linha 3, e não pela
 * posição. Se uma coluna mudar de lugar na extração, a leitura continua
 * correta. A letra em `fallback` (layout padrão da extração) só é usada
 * quando o cabeçalho não é encontrado.
 * ===================================================================== */
const HEADER_ROW = 3; // linha do cabeçalho, numerada como no Excel (1 = primeira linha)

const COLUMN_MAP = {
    NumReserva:  { label: 'Número da Reserva', fallback: 'A',  aliases: ['Nº Res.', 'N Res', 'Nro Res', 'Num Res', 'Numero Res', 'Numero Reserva', 'Número da Reserva', 'Nº Reserva'] },
    Origem:      { label: 'Origem da Reserva', fallback: 'B',  aliases: ['Origem da Reserva', 'Origem'] },
    Prioridade:  { label: 'Prioridade',        fallback: 'C',  aliases: ['Pri', 'Prioridade'] },
    Liberacao:   { label: 'Liberação',         fallback: 'D',  aliases: ['Lib->', 'Lib', 'Liberação', 'Liberado'] },
    StatusOS:    { label: 'Status OS/Orç',     fallback: 'F',  aliases: ['OS/Orç Ofi', 'OS Orc Ofi', 'OS/Orç', 'Status OS'] },
    DocRef:      { label: 'Doc. Referência',   fallback: 'G',  aliases: ['Doc.Ref', 'Doc Ref', 'Documento Referencia'] },
    Codigo:      { label: 'Código da Peça',    fallback: 'H',  aliases: ['Código', 'Código da Peça', 'Cod Peca', 'Cod. Peça'] },
    Qtde:        { label: 'Quantidade',        fallback: 'I',  aliases: ['Qtde.', 'Qtde', 'Qtd', 'Quantidade'] },
    Cliente:     { label: 'Nome do Cliente',   fallback: 'L',  aliases: ['Cliente', 'Nome do Cliente', 'Nome Cliente', 'Cliente Nome'] },
    Telefone:    { label: 'Telefone',          fallback: 'M',  aliases: ['Telefone', 'Fone'] },
    Chassi:      { label: 'Chassi',            fallback: 'U',  aliases: ['CHASSI', 'Chassis'] },
    Placa:       { label: 'Placa',             fallback: 'V',  aliases: ['PLACA'] },
    Consultor:   { label: 'Consultor',         fallback: 'BD', aliases: ['VENDEDOR_NOME', 'Vendedor', 'Nome Vendedor', 'Consultor'] },
    DataReserva: { label: 'Data da Reserva',   fallback: 'AK', aliases: ['DATA_RESERVA', 'Data da Reserva', 'Data Reserva'] },
    Descricao:   { label: 'Descrição da Peça', fallback: 'AM', aliases: ['Descrição', 'Descrição da Peça', 'Descricao Peca', 'Desc. Peça'] },
    // "Dias" (col. E) vem zerada na extração; os dias reais estão em "dias_reservado"
    Dias:        { label: 'Dias Reservado',    fallback: 'AO', aliases: ['dias_reservado', 'Dias Reservado', 'Dias Reserva', 'Dias'] },
    Observacao:  { label: 'Observações',       fallback: 'AU', aliases: ['OBSERVACAO', 'Observação', 'Observações', 'Obs'] },
    Preco:       { label: 'Preço',             fallback: 'AV', aliases: ['PRECO_UNITARIO', 'Preço Unitário', 'Preço', 'Valor Unitário'] },
    TipoReserva: { label: 'Tipo de Reserva',   fallback: 'BB', aliases: ['TIPO_RESERVA', 'Tipo Reserva'] }
};

/* Normaliza textos de cabeçalho: minúsculas, sem acentos, sem pontuação.
 * Ex.: "Nº Res." → "no res" | "OS/Orç Ofi" → "os orc ofi" | "DATA_RESERVA" → "data reserva" */
function normalizeHeader(s) {
    return String(s ?? '')
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}

/* Variações de "Nº" (no / n / nro / num) são tratadas como equivalentes */
function headerKey(s) {
    return normalizeHeader(s).replace(/^(no|n|nro|num|numero)\b/, 'n');
}

/**
 * Resolve o índice (0-based, A = 0) de cada campo a partir do cabeçalho.
 *  1ª passada: nome exato (após normalização), respeitando a ordem dos aliases;
 *  2ª passada: nome parcial (palavra inteira) para os campos não resolvidos;
 *  3ª passada: letra padrão do layout (fallback), se a coluna estiver livre.
 * Uma mesma coluna nunca é atribuída a dois campos.
 */
function resolveColumns(headerRow) {
    const headers = (headerRow || []).map(headerKey);
    const used = new Set();
    const idx = {};
    const origem = {}; // 'cabecalho' | 'fallback' | 'ausente'

    const fields = Object.keys(COLUMN_MAP);

    // 1) Correspondência exata
    for (const f of fields) {
        for (const alias of COLUMN_MAP[f].aliases) {
            const a = headerKey(alias);
            const i = headers.findIndex((h, j) => h && h === a && !used.has(j));
            if (i >= 0) { idx[f] = i; used.add(i); origem[f] = 'cabecalho'; break; }
        }
    }

    // 2) Correspondência parcial por palavra inteira
    for (const f of fields) {
        if (idx[f] !== undefined) continue;
        for (const alias of COLUMN_MAP[f].aliases) {
            const a = headerKey(alias);
            if (a.length < 3) continue;
            const i = headers.findIndex((h, j) => h && !used.has(j) && (` ${h} `).includes(` ${a} `));
            if (i >= 0) { idx[f] = i; used.add(i); origem[f] = 'cabecalho'; break; }
        }
    }

    // 3) Fallback pela letra padrão da extração
    for (const f of fields) {
        if (idx[f] !== undefined) continue;
        const i = XLSX.utils.decode_col(COLUMN_MAP[f].fallback);
        if (!used.has(i) && !headers[i]) {
            idx[f] = i; used.add(i); origem[f] = 'fallback';
        } else {
            idx[f] = -1; origem[f] = 'ausente';
        }
    }

    return { idx, origem };
}

/* Quantos campos conhecidos uma linha contém (usado para validar o cabeçalho) */
function countKnownHeaders(row) {
    const keys = new Set((row || []).map(headerKey).filter(Boolean));
    let n = 0;
    for (const f of Object.keys(COLUMN_MAP)) {
        if (COLUMN_MAP[f].aliases.some(a => keys.has(headerKey(a)))) n++;
    }
    return n;
}

/* Usa a linha 3; se ela não parecer um cabeçalho, procura nas 15 primeiras linhas */
function detectHeaderRow(mat) {
    const preferred = HEADER_ROW - 1;
    if (countKnownHeaders(mat[preferred]) >= 3) return preferred;

    let bestIdx = preferred, best = 0;
    for (let i = 0; i < Math.min(15, mat.length); i++) {
        const n = countKnownHeaders(mat[i]);
        if (n > best) { best = n; bestIdx = i; }
    }
    return bestIdx;
}

function processWorkbook(wb) {
    const name = wb.SheetNames.find(n => {
        const ws = wb.Sheets[n];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        return (range.e.r - range.s.r) >= 1;
    }) || wb.SheetNames[0];

    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) { showError('A planilha selecionada está vazia.'); return; }

    // Lê sempre a partir de A1, para que o índice da matriz corresponda
    // exatamente à linha/coluna do Excel (mat[2] = linha 3, coluna 0 = A).
    const full = XLSX.utils.decode_range(ws['!ref']);
    const mat = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        blankrows: true,
        defval: '',
        raw: true,
        range: { s: { r: 0, c: 0 }, e: full.e }
    });
    if (!mat || !mat.length) { showError('Nenhuma linha de dados encontrada na planilha.'); return; }

    const headerIdx = detectHeaderRow(mat);
    const { idx, origem } = resolveColumns(mat[headerIdx]);

    // Linhas de dados: tudo abaixo do cabeçalho, ignorando linhas totalmente vazias
    const dataRows = mat.slice(headerIdx + 1)
        .filter(r => Array.isArray(r) && r.some(c => String(c ?? '').trim() !== ''));

    logColumnMapping(mat[headerIdx], idx, origem, headerIdx);

    const naoEncontrados = Object.keys(origem).filter(f => origem[f] !== 'cabecalho').map(f => COLUMN_MAP[f].label);
    if (naoEncontrados.length) {
        showWarn(`Colunas não localizadas pelo cabeçalho da linha ${headerIdx + 1}: ${naoEncontrados.join(', ')}. Foi usada a posição padrão da extração (quando disponível).`);
    }

    parseRows(idx, dataRows);
}

/* Exibe no console (F12) qual coluna do Excel foi usada para cada campo */
function logColumnMapping(headerRow, idx, origem, headerIdx) {
    if (!console || !console.table) return;
    const tbl = {};
    for (const f of Object.keys(COLUMN_MAP)) {
        const i = idx[f];
        tbl[COLUMN_MAP[f].label] = {
            Coluna: i >= 0 ? XLSX.utils.encode_col(i) : '—',
            Cabecalho: i >= 0 ? String((headerRow || [])[i] ?? '') : '—',
            Origem: origem[f]
        };
    }
    console.info(`[Reservas] Cabeçalho lido na linha ${headerIdx + 1}. Mapeamento de colunas:`);
    console.table(tbl);
}

function parseRows(idx, data) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    allData = data.map(r => {
        const raw = (f) => (idx[f] >= 0 && r[idx[f]] !== undefined && r[idx[f]] !== null) ? r[idx[f]] : '';
        const txt = (f, vazio = '—') => {
            const s = String(raw(f)).trim();
            return s || vazio;
        };

        // Data da reserva
        const dataRaw = raw('DataReserva');
        const dataObj = (dataRaw === '' ? null : parseDateExcel(dataRaw));
        const dataTxt = typeof dataRaw === 'number'
            ? (dataObj ? dataObj.toLocaleDateString('pt-BR') : '—')
            : (String(dataRaw).trim() || '—');

        // Dias em reserva: usa a coluna; se não existir, calcula pela data da reserva
        let diasNum = toNumberBR(raw('Dias'));
        if (idx.Dias < 0 && dataObj && !isNaN(+dataObj)) {
            const d0 = new Date(dataObj); d0.setHours(0, 0, 0, 0);
            diasNum = Math.max(0, Math.round((hoje - d0) / 86400000));
        }

        const libRaw = txt('Liberacao', '');
        const qtde = toNumberBR(raw('Qtde')) || 1;
        const preco = toNumberBR(raw('Preco'));
        const consultor = txt('Consultor', 'Não informado');

        return {
            NumReserva:   txt('NumReserva'),
            Origem:       normalizeOrigem(txt('Origem', '')),
            Prioridade:   txt('Prioridade', '0'),
            Liberacao:    normalizeLiberacao(libRaw),
            LiberacaoRaw: libRaw || '—',
            Dias:         diasNum,
            FaixaDias:    mapFaixaDias(diasNum),
            StatusOS:     txt('StatusOS'),
            DocRef:       txt('DocRef'),
            Codigo:       txt('Codigo'),
            Descricao:    txt('Descricao'),
            Qtde:         qtde,
            Preco:        preco,
            ValorTotal:   preco * qtde,
            Cliente:      txt('Cliente', 'Não informado'),
            Telefone:     txt('Telefone'),
            Chassi:       txt('Chassi'),
            Placa:        txt('Placa'),
            Consultor:    consultor,
            DataReserva:  dataTxt,
            MesAno:       monthYearFromCell(dataRaw),
            Observacao:   txt('Observacao'),
            TipoReserva:  txt('TipoReserva'),
            Vendedor:     consultor
        };
    });

    populateFilterOptions();
    applyFilters();
}

function populateFilterOptions() {
    // Meses em ordem cronológica
    const orderMes = MONTH_NAMES.map(m => m.toLowerCase());
    const mesesUnicos = [...new Set(allData.map(x => x.MesAno).filter(m => m && m !== 'Não informado'))]
        .sort((a, b) => {
            const pa = a.split(' '), pb = b.split(' ');
            const ma = orderMes.indexOf(pa[0].toLowerCase()), mb = orderMes.indexOf(pb[0].toLowerCase());
            const ya = parseInt(pa[1], 10), yb = parseInt(pb[1], 10);
            if (ya !== yb) return ya - yb;
            return ma - mb;
        });

    pillInstances.Mes.setOptions(mesesUnicos, { keepOrder: true });
    pillInstances.Origem.setOptions(allData.map(x => x.Origem));
    pillInstances.Vendedor.setOptions(allData.map(x => x.Vendedor));
    pillInstances.Codigo.setOptions(allData.map(x => x.Codigo));
    pillInstances.Prioridade.setOptions(allData.map(x => x.Prioridade));
    pillInstances.Liberacao.setOptions(['Liberada', 'Não Liberada', 'Pendente / Não Informado'], { keepOrder: true });
    pillInstances.StatusOS.setOptions(allData.map(x => x.StatusOS).filter(s => s && s !== '—'));
    pillInstances.FaixaDias.setOptions([
        'de 0 a 30 dias',
        'de 31 a 90 dias',
        'de 91 a 180 dias',
        'de 181 a 365 dias',
        'acima de 365 dias'
    ], { keepOrder: true });
}

/* ===== Busca e Limpar Filtros ===== */
const searchInput = document.getElementById('q');
if (searchInput) {
    searchInput.addEventListener('input', e => {
        state.search = e.target.value.toLowerCase().trim();
        withStableScroll(() => applyFilters());
    });
}

const clearAllBtn = document.getElementById('clearAll');
if (clearAllBtn) {
    clearAllBtn.addEventListener('click', () => {
        withStableScroll(() => {
            state.search = '';
            if (searchInput) searchInput.value = '';
            state.Mes = [];
            state.Origem = [];
            state.Vendedor = [];
            state.Codigo = [];
            state.Prioridade = [];
            state.Liberacao = [];
            state.StatusOS = [];
            state.FaixaDias = [];
            barFilter = null;
            Object.values(pillInstances).forEach(p => {
                p.setter([]);
                p.render();
                p.sync();
            });
            applyFilters();
        });
    });
}

/* ===== Aplicação dos Filtros ===== */
function applyFilters() {
    filtered = allData.filter(r => {
        const mesOk = state.Mes.length === 0 || state.Mes.includes(r.MesAno);
        const origemOk = state.Origem.length === 0 || state.Origem.includes(r.Origem);
        const vendOk = state.Vendedor.length === 0 || state.Vendedor.includes(r.Vendedor);
        const codOk = state.Codigo.length === 0 || state.Codigo.includes(r.Codigo);
        const priOk = state.Prioridade.length === 0 || state.Prioridade.includes(r.Prioridade);
        const libOk = state.Liberacao.length === 0 || state.Liberacao.includes(r.Liberacao);
        const statusOk = state.StatusOS.length === 0 || state.StatusOS.includes(r.StatusOS);
        const diasOk = state.FaixaDias.length === 0 || state.FaixaDias.includes(r.FaixaDias);

        const q = state.search;
        const qOk = !q || [
            r.NumReserva,
            r.Codigo,
            r.Descricao,
            r.Cliente,
            r.Vendedor,
            r.Consultor,
            r.Chassi,
            r.Placa,
            r.Origem,
            r.StatusOS,
            r.DocRef,
            r.Telefone,
            r.Observacao
        ].some(val => (val || '').toString().toLowerCase().includes(q));

        return mesOk && origemOk && vendOk && codOk && priOk && libOk && statusOk && diasOk && qOk;
    });

    currentPage = 1;
    renderInsights(filtered);
    renderKpiGroups();
    renderCharts(filtered);
    renderTableSection();
}

/* ===== Filtro de Barra dos Gráficos ===== */
function filterByBar(rows, filter) {
    if (!filter || !filter.label) return rows;
    const { type, label } = filter;

    return rows.filter(r => {
        if (type === 'periodo') return r.MesAno === label;
        if (type === 'topPecas') return r.Codigo === label;
        if (type === 'vendedor') return (r.Vendedor || 'Não informado') === label;
        if (type === 'origem') return r.Origem === label;
        if (type === 'status') return r.StatusOS === label;
        if (type === 'dias') return r.FaixaDias === label;
        if (type === 'prioridade') return String(r.Prioridade) === label;
        return true;
    });
}

function onChartBarClick(evt, elements, chartInstance, chartType) {
    if (!elements || !elements.length) return;
    const idx = elements[0].index;
    const label = chartInstance.data.labels[idx];
    if (!label) return;

    barFilter = { type: chartType, label: label };
    tableLoaded = true;
    currentPage = 1;
    renderTableSection();

    const tableEl = document.getElementById('tableToggleBar');
    if (tableEl) {
        tableEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

/* ===== Storytelling / Insights ===== */
function renderInsights(rows) {
    const grid = document.getElementById('insightsGrid');
    if (!grid) return;
    grid.innerHTML = '';
    const n = rows.length;

    if (!n) {
        grid.innerHTML = `
            <div class="insight-card">
                <div class="insight-ico"><i class="fa-solid fa-circle-info"></i></div>
                <div class="insight-body">
                    <div class="insight-head">Sem dados com os filtros atuais</div>
                    <div class="insight-txt" style="color: #64748b;">Ajuste ou limpe os filtros para visualizar os principais insights do dashboard.</div>
                </div>
            </div>`;
        return;
    }

    const totalPecas = rows.reduce((s, r) => s + (+r.Qtde || 0), 0);
    const mediaPecas = n ? (totalPecas / n) : 0;

    // Peça mais reservada
    const pecaMap = {};
    rows.forEach(r => {
        const k = r.Codigo || '—';
        pecaMap[k] = (pecaMap[k] || 0) + (+r.Qtde || 0);
    });
    const topPeca = Object.entries(pecaMap).sort((a, b) => b[1] - a[1])[0];

    // Vendedor destaque
    const vendMap = {};
    rows.forEach(r => {
        const v = r.Vendedor || 'Não informado';
        vendMap[v] = (vendMap[v] || 0) + (+r.Qtde || 0);
    });
    const topVend = Object.entries(vendMap).sort((a, b) => b[1] - a[1])[0];

    // Origem predominante
    const origemMap = {};
    rows.forEach(r => {
        const o = r.Origem || 'Não informado';
        origemMap[o] = (origemMap[o] || 0) + 1;
    });
    const topOrigem = Object.entries(origemMap).sort((a, b) => b[1] - a[1])[0];
    const pctOrigem = topOrigem ? ((topOrigem[1] / n) * 100).toFixed(1).replace('.', ',') : '0';

    // Liberações
    const liberadas = rows.filter(r => r.Liberacao === 'Liberada').length;
    const naoLiberadas = rows.filter(r => r.Liberacao === 'Não Liberada').length;
    const pctLib = ((liberadas / n) * 100).toFixed(1).replace('.', ',');

    // Clientes únicos
    const clientesSet = new Set(rows.map(r => (r.Cliente || '').trim().toUpperCase()).filter(c => c && c !== 'NÃO INFORMADO' && c !== '—'));

    const cards = [
        {
            icon: 'fa-tags',
            head: 'Peça Mais Reservada',
            text: `Código <span class="highlight">${topPeca ? topPeca[0] : '—'}</span> lidera com <span class="highlight">${NUM(topPeca ? topPeca[1] : 0)}</span> peça(s) reservada(s).`
        },
        {
            icon: 'fa-user-tie',
            head: 'Vendedor Líder',
            text: `<span class="highlight">${topVend ? topVend[0] : '—'}</span> acumula <span class="highlight">${NUM(topVend ? topVend[1] : 0)}</span> peças em reservas.`
        },
        {
            icon: 'fa-diagram-project',
            head: 'Origem Predominante',
            text: `<span class="highlight">${topOrigem ? topOrigem[0] : '—'}</span> representa <span class="highlight">${pctOrigem}%</span> do total (${NUM(topOrigem ? topOrigem[1] : 0)} reservas).`
        },
        {
            icon: 'fa-shield-halved',
            head: 'Taxa de Liberação',
            text: `<span class="highlight">${pctLib}%</span> das reservas liberadas (${NUM(liberadas)} de ${NUM(n)}).`
        },
        {
            icon: 'fa-boxes-stacked',
            head: 'Média de Peças / Reserva',
            text: `Média de <span class="highlight">${mediaPecas.toFixed(1).replace('.', ',')}</span> peças por registro de reserva.`
        },
        {
            icon: 'fa-users',
            head: 'Clientes Atendidos',
            text: `<span class="highlight">${NUM(clientesSet.size)}</span> clientes/solicitantes distintos registrados.`
        }
    ];

    grid.innerHTML = cards.map(c => `
        <div class="insight-card">
            <div class="insight-ico"><i class="fa-solid ${c.icon}"></i></div>
            <div class="insight-body">
                <div class="insight-head">${c.head}</div>
                <div class="insight-txt">${c.text}</div>
            </div>
        </div>
    `).join('');
}

/* ===== KPIs ===== */
function calcKPIs(rows) {
    const totalReservas = rows.length;
    const totalPecas = rows.reduce((s, r) => s + (+r.Qtde || 0), 0);
    const mediaPecas = totalReservas ? totalPecas / totalReservas : 0;

    // Peça mais reservada
    const pecaMap = {};
    rows.forEach(r => {
        const k = r.Codigo || '—';
        pecaMap[k] = (pecaMap[k] || 0) + (+r.Qtde || 0);
    });
    const topPeca = Object.entries(pecaMap).sort((a, b) => b[1] - a[1])[0];

    // Vendedor líder
    const vendMap = {};
    rows.forEach(r => {
        const v = r.Vendedor || 'Não informado';
        vendMap[v] = (vendMap[v] || 0) + (+r.Qtde || 0);
    });
    const topVend = Object.entries(vendMap).sort((a, b) => b[1] - a[1])[0];

    // Pendentes / Não Liberadas
    const naoLiberadas = rows.filter(r => r.Liberacao === 'Não Liberada').length;
    const pendentes = rows.filter(r => r.Liberacao.includes('Pendente')).length;

    // Dias médios
    const diasVals = rows.map(r => +r.Dias || 0).filter(d => d > 0);
    const diasMedios = diasVals.length ? diasVals.reduce((a, b) => a + b, 0) / diasVals.length : 0;

    return {
        totalReservas,
        totalPecas,
        mediaPecas,
        itemMaisReservado: topPeca ? `${topPeca[0]} (${NUM(topPeca[1])} un.)` : '—',
        vendedorLider: topVend ? `${topVend[0]} (${NUM(topVend[1])} un.)` : '—',
        naoLiberadas,
        pendentes,
        diasMedios
    };
}

function renderKpiGroups() {
    const k = calcKPIs(filtered);
    const holder = document.getElementById('kpiGroups');
    if (!holder) return;

    holder.innerHTML = `
        <div class="kpi-group">
            <div class="kpi-grid">
                <div class="kpi">
                    <div class="lab">Total de Reservas</div>
                    <div class="val">${NUM(k.totalReservas)}</div>
                </div>
                <div class="kpi">
                    <div class="lab">Total Peças Reservadas</div>
                    <div class="val">${NUM(k.totalPecas)}</div>
                </div>
                <div class="kpi">
                    <div class="lab">Média Peças / Reserva</div>
                    <div class="val">${DEC(k.mediaPecas)}</div>
                </div>
                <div class="kpi">
                    <div class="lab">Peça Mais Reservada</div>
                    <div class="val sub">${k.itemMaisReservado}</div>
                </div>
                <div class="kpi">
                    <div class="lab">Vendedor Líder</div>
                    <div class="val sub">${k.vendedorLider}</div>
                </div>
                <div class="kpi">
                    <div class="lab">Não Liberadas / Pendentes</div>
                    <div class="val" style="color: ${k.naoLiberadas > 0 ? 'var(--danger)' : 'var(--primary)'};">
                        ${NUM(k.naoLiberadas)} <span style="font-size: 0.9rem; color: #64748b; font-weight: 600;">(${NUM(k.pendentes)} pend.)</span>
                    </div>
                </div>
                <div class="kpi">
                    <div class="lab">Média de Dias em Reserva</div>
                    <div class="val">${DEC(k.diasMedios)} dias</div>
                </div>
            </div>
        </div>
    `;
}

/* ===== Gráficos (Chart.js) ===== */
function renderCharts(rows) {
    const commonOpts = {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
            datalabels: {}
        }
    };

    /* 1. Reservas por Período (Mês/Ano) */
    const statsByMonth = {};
    rows.forEach(r => {
        const k = r.MesAno || 'Não informado';
        if (!statsByMonth[k]) statsByMonth[k] = { reservas: 0, pecas: 0 };
        statsByMonth[k].reservas += 1;
        statsByMonth[k].pecas += (+r.Qtde || 0);
    });

    const orderMes = MONTH_NAMES.map(m => m.toLowerCase());
    const mesesSorted = Object.entries(statsByMonth).sort((a, b) => {
        const pa = a[0].split(' '), pb = b[0].split(' ');
        if (pa[0] === 'Não' || pb[0] === 'Não') return 0;
        const ma = orderMes.indexOf(pa[0].toLowerCase()), mb = orderMes.indexOf(pb[0].toLowerCase());
        const ya = parseInt(pa[1], 10), yb = parseInt(pb[1], 10);
        if (ya !== yb) return ya - yb;
        return ma - mb;
    });

    if (charts.periodo) charts.periodo.destroy();
    const periodoCtx = document.getElementById('periodoChart')?.getContext('2d');
    if (periodoCtx) {
        charts.periodo = new Chart(periodoCtx, {
            type: 'bar',
            data: {
                labels: mesesSorted.map(x => x[0]),
                datasets: [
                    {
                        label: 'Total de Peças',
                        data: mesesSorted.map(x => x[1].pecas),
                        backgroundColor: '#3b82f6',
                        borderRadius: 6,
                        barPercentage: 0.6,
                        categoryPercentage: 0.8
                    },
                    {
                        label: 'Qtd. Reservas',
                        data: mesesSorted.map(x => x[1].reservas),
                        backgroundColor: '#10b981',
                        borderRadius: 6,
                        barPercentage: 0.6,
                        categoryPercentage: 0.8
                    }
                ]
            },
            options: {
                ...commonOpts,
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.periodo, 'periodo'),
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    },
                    y: {
                        grace: '15%',
                        grid: { color: '#f1f5f9' },
                        ticks: { font: { size: 11 }, color: '#64748b' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    title: {
                        display: true,
                        text: 'Histórico de Reservas e Volume de Peças por Mês',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'top',
                        anchor: 'end',
                        color: '#1e3a8a',
                        font: { weight: 'bold', size: 10 },
                        formatter: v => NUM(v)
                    },
                    tooltip: {
                        callbacks: {
                            label: context => {
                                const name = context.label;
                                const s = statsByMonth[name] || { reservas: 0, pecas: 0 };
                                if (context.datasetIndex === 0) {
                                    return `Total de Peças: ${NUM(s.pecas)}`;
                                }
                                return `Qtd. de Reservas: ${NUM(s.reservas)}`;
                            }
                        }
                    }
                }
            }
        });
    }

    /* 2. Top 20 Peças Mais Reservadas (Quantidade) */
    const statsByItem = {};
    rows.forEach(r => {
        const k = r.Codigo || '—';
        if (!statsByItem[k]) statsByItem[k] = { pecas: 0, reservas: 0 };
        statsByItem[k].pecas += (+r.Qtde || 0);
        statsByItem[k].reservas += 1;
    });

    const topPecasArr = Object.entries(statsByItem)
        .sort((a, b) => b[1].pecas - a[1].pecas)
        .slice(0, 20);

    const totalPecasTop = rows.reduce((s, r) => s + (+r.Qtde || 0), 0) || 1;

    if (charts.topPecas) charts.topPecas.destroy();
    const topPecasCtx = document.getElementById('topPecasChart')?.getContext('2d');
    if (topPecasCtx) {
        charts.topPecas = new Chart(topPecasCtx, {
            type: 'bar',
            data: {
                labels: topPecasArr.map(x => x[0]),
                datasets: [{
                    label: 'Quantidade Reservada',
                    data: topPecasArr.map(x => x[1].pecas),
                    backgroundColor: '#8b5cf6',
                    borderRadius: 6,
                    barPercentage: 0.75,
                    categoryPercentage: 0.85
                }]
            },
            options: {
                ...commonOpts,
                indexAxis: 'y',
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.topPecas, 'topPecas'),
                scales: {
                    x: {
                        grace: '20%',
                        grid: { color: '#f1f5f9' },
                        ticks: { font: { size: 11 }, color: '#64748b' }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Top 20 Peças Mais Reservadas (Quantidade)',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'end',
                        anchor: 'end',
                        offset: 6,
                        color: '#6d28d9',
                        font: { weight: 'bold', size: 11 },
                        formatter: (v) => {
                            const pct = ((v / totalPecasTop) * 100).toFixed(1).replace('.', ',');
                            return `${NUM(v)} (${pct}%)`;
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: context => {
                                const cod = context.label;
                                const s = statsByItem[cod] || { pecas: 0, reservas: 0 };
                                const pct = ((s.pecas / totalPecasTop) * 100).toFixed(1).replace('.', ',');
                                return [
                                    `Código: ${cod}`,
                                    `Quantidade Total: ${NUM(s.pecas)} peças (${pct}%)`,
                                    `Total de Reservas: ${NUM(s.reservas)} pedidos`
                                ];
                            }
                        }
                    }
                }
            }
        });
    }

    /* 3. Top 15 Vendedores por Quantidade de Peças */
    const statsByVend = {};
    rows.forEach(r => {
        const k = r.Vendedor || 'Não informado';
        if (!statsByVend[k]) statsByVend[k] = { pecas: 0, reservas: 0 };
        statsByVend[k].pecas += (+r.Qtde || 0);
        statsByVend[k].reservas += 1;
    });

    const topVendArr = Object.entries(statsByVend)
        .sort((a, b) => b[1].pecas - a[1].pecas)
        .slice(0, 15);

    if (charts.vendedor) charts.vendedor.destroy();
    const vendedorCtx = document.getElementById('vendedorChart')?.getContext('2d');
    if (vendedorCtx) {
        charts.vendedor = new Chart(vendedorCtx, {
            type: 'bar',
            data: {
                labels: topVendArr.map(x => x[0]),
                datasets: [{
                    label: 'Quantidade de Peças',
                    data: topVendArr.map(x => x[1].pecas),
                    backgroundColor: '#3b82f6',
                    borderRadius: 6,
                    barPercentage: 0.75,
                    categoryPercentage: 0.85
                }]
            },
            options: {
                ...commonOpts,
                indexAxis: 'y',
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.vendedor, 'vendedor'),
                scales: {
                    x: {
                        grace: '20%',
                        grid: { color: '#f1f5f9' },
                        ticks: { font: { size: 11 }, color: '#64748b' }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Top 15 Vendedores por Volume de Peças',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'end',
                        anchor: 'end',
                        offset: 6,
                        color: '#1d4ed8',
                        font: { weight: 'bold', size: 11 },
                        formatter: v => NUM(v)
                    },
                    tooltip: {
                        callbacks: {
                            label: context => {
                                const name = context.label;
                                const s = statsByVend[name] || { pecas: 0, reservas: 0 };
                                return [
                                    `Vendedor: ${name}`,
                                    `Volume de Peças: ${NUM(s.pecas)}`,
                                    `Total de Reservas: ${NUM(s.reservas)}`
                                ];
                            }
                        }
                    }
                }
            }
        });
    }

    /* 4. Distribuição por Origem da Reserva */
    const statsByOrigem = {};
    rows.forEach(r => {
        const k = r.Origem || 'Não informado';
        statsByOrigem[k] = (statsByOrigem[k] || 0) + 1;
    });

    const origemArr = Object.entries(statsByOrigem).sort((a, b) => b[1] - a[1]);
    const totalOrigem = rows.length || 1;

    if (charts.origem) charts.origem.destroy();
    const origemCtx = document.getElementById('origemChart')?.getContext('2d');
    if (origemCtx) {
        charts.origem = new Chart(origemCtx, {
            type: 'bar',
            data: {
                labels: origemArr.map(x => x[0]),
                datasets: [{
                    label: 'Reservas',
                    data: origemArr.map(x => x[1]),
                    backgroundColor: ['#1e40af', '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#64748b'],
                    borderRadius: 8
                }]
            },
            options: {
                ...commonOpts,
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.origem, 'origem'),
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    },
                    y: {
                        grace: '20%',
                        grid: { color: '#f1f5f9' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Distribuição por Origem da Reserva',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'end',
                        anchor: 'end',
                        offset: 4,
                        color: '#1e3a8a',
                        font: { weight: 'bold', size: 11 },
                        formatter: (v) => {
                            const pct = ((v / totalOrigem) * 100).toFixed(1).replace('.', ',');
                            return `${NUM(v)} (${pct}%)`;
                        }
                    }
                }
            }
        });
    }

    /* 5. Distribuição por Status OS / Orçamento */
    const statsByStatus = {};
    rows.forEach(r => {
        const k = (r.StatusOS && r.StatusOS !== '—') ? r.StatusOS : 'Sem OS/Orçamento';
        statsByStatus[k] = (statsByStatus[k] || 0) + 1;
    });

    const statusArr = Object.entries(statsByStatus).sort((a, b) => b[1] - a[1]);
    const totalStatus = rows.length || 1;

    if (charts.status) charts.status.destroy();
    const statusCtx = document.getElementById('statusChart')?.getContext('2d');
    if (statusCtx) {
        charts.status = new Chart(statusCtx, {
            type: 'bar',
            data: {
                labels: statusArr.map(x => x[0]),
                datasets: [{
                    label: 'Reservas',
                    data: statusArr.map(x => x[1]),
                    backgroundColor: ['#64748b', '#ef4444', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6'],
                    borderRadius: 8
                }]
            },
            options: {
                ...commonOpts,
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.status, 'status'),
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    },
                    y: {
                        grace: '20%',
                        grid: { color: '#f1f5f9' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Distribuição por Status de OS / Orçamento',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'end',
                        anchor: 'end',
                        offset: 4,
                        color: '#334155',
                        font: { weight: 'bold', size: 11 },
                        formatter: (v) => {
                            const pct = ((v / totalStatus) * 100).toFixed(1).replace('.', ',');
                            return `${NUM(v)} (${pct}%)`;
                        }
                    }
                }
            }
        });
    }

    /* 6. Faixa de Dias em Reserva */
    const faixasDias = {
        'de 0 a 30 dias': 0,
        'de 31 a 90 dias': 0,
        'de 91 a 180 dias': 0,
        'de 181 a 365 dias': 0,
        'acima de 365 dias': 0
    };

    rows.forEach(r => {
        const faixa = r.FaixaDias || mapFaixaDias(r.Dias);
        if (faixasDias[faixa] !== undefined) {
            faixasDias[faixa]++;
        }
    });

    const labelsDias = Object.keys(faixasDias);
    const valoresDias = labelsDias.map(k => faixasDias[k]);
    const coresDias = ['#10b981', '#3b82f6', '#f59e0b', '#f97316', '#dc2626'];
    const totalDias = valoresDias.reduce((a, b) => a + b, 0) || 1;

    if (charts.dias) charts.dias.destroy();
    const diasCtx = document.getElementById('diasChart')?.getContext('2d');
    if (diasCtx) {
        charts.dias = new Chart(diasCtx, {
            type: 'bar',
            data: {
                labels: labelsDias,
                datasets: [{
                    label: 'Reservas',
                    data: valoresDias,
                    backgroundColor: coresDias,
                    borderRadius: 8,
                    barPercentage: 0.65,
                    categoryPercentage: 0.8
                }]
            },
            options: {
                ...commonOpts,
                indexAxis: 'y',
                onClick: (evt, elements) => onChartBarClick(evt, elements, charts.dias, 'dias'),
                scales: {
                    x: {
                        grace: '25%',
                        grid: { color: '#f1f5f9' },
                        ticks: { font: { size: 11 }, color: '#64748b' }
                    },
                    y: {
                        grid: { display: false },
                        ticks: { font: { size: 11, weight: '600' }, color: '#334155' }
                    }
                },
                plugins: {
                    ...commonOpts.plugins,
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Tempo em Reserva (Faixas de Idade das Reservas)',
                        font: { size: 14, weight: 'bold' },
                        color: '#1e293b'
                    },
                    datalabels: {
                        display: true,
                        align: 'end',
                        anchor: 'end',
                        offset: 6,
                        color: '#1e293b',
                        font: { weight: 'bold', size: 11 },
                        formatter: (v) => {
                            const pct = ((v / totalDias) * 100).toFixed(1).replace('.', ',');
                            return `${NUM(v)} reservas (${pct}%)`;
                        }
                    }
                }
            }
        });
    }
}

/* ===== Tabela sob Demanda & Paginação ===== */
function renderTableSection() {
    const chipEl = document.getElementById('barFilterChip');
    const clearBtn = document.getElementById('clearBarFilterBtn');
    const toggleBtn = document.getElementById('toggleTableBtn');
    const placeholder = document.getElementById('tablePlaceholder');
    const gridWrap = document.getElementById('gridWrap');

    let currentRows = barFilter ? filterByBar(filtered, barFilter) : filtered;

    if (barFilter && barFilter.label) {
        chipEl.style.display = 'inline-flex';
        chipEl.innerHTML = `<i class="fa-solid fa-filter"></i> Filtrado por: <strong>${barFilter.label}</strong> (${NUM(currentRows.length)} reservas)`;
        clearBtn.style.display = 'inline-flex';
    } else {
        chipEl.style.display = 'none';
        clearBtn.style.display = 'none';
    }

    if (!tableLoaded) {
        placeholder.style.display = 'block';
        gridWrap.style.display = 'none';
        toggleBtn.innerHTML = `<i class="fa-solid fa-folder-open"></i> Carregar Reservas (${NUM(currentRows.length)})`;
    } else {
        placeholder.style.display = 'none';
        gridWrap.style.display = 'block';
        toggleBtn.innerHTML = `<i class="fa-solid fa-folder-minus"></i> Ocultar Reservas`;

        const sorted = sortRows([...currentRows]);
        renderTableSlice(sorted);
    }
}

function renderTableSlice(rows) {
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, total);
    const slice = rows.slice(startIdx, endIdx);

    renderTable(slice);

    const pagEl = document.getElementById('tablePagination');
    if (!pagEl) return;

    if (total === 0) {
        pagEl.innerHTML = `<div class="pagination-info">Nenhuma reserva encontrada para exibição.</div>`;
        return;
    }

    pagEl.innerHTML = `
        <div class="pagination-info">Exibindo <strong>${NUM(startIdx + 1)}–${NUM(endIdx)}</strong> de <strong>${NUM(total)}</strong> reservas</div>
        <div class="pagination-controls">
            <button class="pagination-btn" id="pagFirst" ${currentPage === 1 ? 'disabled' : ''}><i class="fa-solid fa-angles-left"></i></button>
            <button class="pagination-btn" id="pagPrev" ${currentPage === 1 ? 'disabled' : ''}><i class="fa-solid fa-angle-left"></i></button>
            <span class="pagination-info" style="margin: 0 6px;">Pág. ${currentPage} de ${totalPages}</span>
            <button class="pagination-btn" id="pagNext" ${currentPage === totalPages ? 'disabled' : ''}><i class="fa-solid fa-angle-right"></i></button>
            <button class="pagination-btn" id="pagLast" ${currentPage === totalPages ? 'disabled' : ''}><i class="fa-solid fa-angles-right"></i></button>
        </div>
    `;

    document.getElementById('pagFirst')?.addEventListener('click', () => { currentPage = 1; renderTableSlice(rows); });
    document.getElementById('pagPrev')?.addEventListener('click', () => { currentPage--; renderTableSlice(rows); });
    document.getElementById('pagNext')?.addEventListener('click', () => { currentPage++; renderTableSlice(rows); });
    document.getElementById('pagLast')?.addEventListener('click', () => { currentPage = totalPages; renderTableSlice(rows); });
}

function renderTable(rows) {
    const tb = document.querySelector('#grid tbody');
    if (!tb) return;
    tb.innerHTML = '';

    for (const r of rows) {
        const tr = document.createElement('tr');

        // Badge de liberação
        let libBadge = `<span class="badge badge-lib-pend">Pendente</span>`;
        if (r.Liberacao === 'Liberada' || r.LiberacaoRaw === 'S') {
            libBadge = `<span class="badge badge-lib-s"><i class="fa-solid fa-check"></i> Liberada (S)</span>`;
        } else if (r.Liberacao === 'Não Liberada' || r.LiberacaoRaw === 'N') {
            libBadge = `<span class="badge badge-lib-n"><i class="fa-solid fa-xmark"></i> Bloqueada (N)</span>`;
        }

        tr.innerHTML = `
            <td><strong>#${r.NumReserva}</strong></td>
            <td>${r.DataReserva || '—'}</td>
            <td>${r.Cliente || '—'}</td>
            <td class="cell-vendedor" title="Filtrar por este consultor">${r.Consultor || r.Vendedor || '—'}</td>
            <td class="cell-codigo" title="Filtrar por esta peça">${r.Codigo}</td>
            <td class="cell-ellipsis" title="${escAttr(r.Descricao)}">${r.Descricao || '—'}</td>
            <td><strong>${NUM(r.Qtde)}</strong></td>
            <td>${BRL(r.Preco)}</td>
            <td>${NUM(r.Dias)}</td>
            <td class="cell-ellipsis" title="${escAttr(r.Observacao)}">${r.Observacao || '—'}</td>
            <td><span class="badge badge-origem">${r.Origem}</span></td>
            <td><span class="badge badge-pri">Pri ${r.Prioridade}</span></td>
            <td>${libBadge}</td>
            <td>${r.StatusOS || '—'}</td>
        `;
        tb.appendChild(tr);
    }
}

/* ===== Ordenação da Tabela ===== */
const collator = new Intl.Collator('pt-BR', { numeric: true, sensitivity: 'base' });

function sortRows(rows) {
    if (!sortState.key) return rows;
    const key = sortState.key, type = sortState.type, dir = sortState.dir === 'asc' ? 1 : -1;
    return rows.sort((a, b) => {
        let va = a[key], vb = b[key];
        if (type === 'num') {
            va = Number(va) || 0;
            vb = Number(vb) || 0;
            return dir * (va - vb);
        } else {
            return dir * collator.compare(String(va || ''), String(vb || ''));
        }
    });
}

function initSorting() {
    const ths = document.querySelectorAll('#grid thead th.sortable');
    ths.forEach(th => {
        th.addEventListener('click', () => withStableScroll(() => {
            const key = th.dataset.key, type = th.dataset.type;
            if (sortState.key === key) {
                sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
            } else {
                sortState.key = key;
                sortState.type = type;
                sortState.dir = 'desc';
            }
            ths.forEach(t => {
                t.classList.remove('active');
                const icon = t.querySelector('.sort');
                if (icon) icon.className = 'fa-solid fa-sort sort';
            });
            th.classList.add('active');
            const icon = th.querySelector('.sort');
            if (icon) {
                icon.className = 'fa-solid ' + (sortState.dir === 'asc' ? 'fa-sort-up' : 'fa-sort-down') + ' sort';
            }
            renderTableSection();
        }));
    });
}
initSorting();

/* Event listeners da Tabela sob Demanda */
document.getElementById('toggleTableBtn')?.addEventListener('click', (e) => {
    e.stopPropagation();
    tableLoaded = !tableLoaded;
    renderTableSection();
});

document.getElementById('tableToggleBar')?.addEventListener('click', (e) => {
    if (e.target.closest('button')) return;
    tableLoaded = !tableLoaded;
    renderTableSection();
});

document.getElementById('clearBarFilterBtn')?.addEventListener('click', (e) => {
    e.stopPropagation();
    barFilter = null;
    currentPage = 1;
    renderTableSection();
});

function exportToExcel() {
    if (!window.XLSX) {
        showError('Biblioteca XLSX não carregada. Verifique sua conexão com a internet.');
        return;
    }

    const baseRows = barFilter ? filterByBar(filtered, barFilter) : filtered;
    const rowsToExport = sortRows([...baseRows]);

    if (!rowsToExport.length) {
        showWarn('Não há reservas para exportar com os filtros atuais.');
        return;
    }

    const headers = [
        'Nº Reserva',
        'Data da Reserva',
        'Cliente',
        'Consultor',
        'Código da Peça',
        'Descrição da Peça',
        'Qtde',
        'Preço Unitário',
        'Valor Total',
        'Dias',
        'Observações',
        'Origem',
        'Prioridade',
        'Liberação',
        'Status OS'
    ];

    const aoa = [headers];
    for (const r of rowsToExport) {
        const qtde = Number(r.Qtde) || 0;
        const preco = Number(r.Preco) || 0;
        const total = Number(r.ValorTotal) || (preco * qtde);
        const dias = Number(r.Dias) || 0;

        aoa.push([
            r.NumReserva || '',
            r.DataReserva || '',
            r.Cliente || '',
            r.Consultor || r.Vendedor || '',
            r.Codigo || '',
            r.Descricao || '',
            qtde,
            preco,
            total,
            dias,
            r.Observacao === '—' ? '' : (r.Observacao || ''),
            r.Origem || '',
            r.Prioridade || '',
            r.Liberacao || '',
            r.StatusOS === '—' ? '' : (r.StatusOS || '')
        ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Formata colunas monetárias e ajusta largura automática das colunas
    const colWidths = headers.map((h, cIdx) => {
        let maxLen = h.length;
        for (let rIdx = 1; rIdx < aoa.length; rIdx++) {
            const val = aoa[rIdx][cIdx];
            const len = String(val ?? '').length;
            if (len > maxLen) maxLen = len;
            if ((cIdx === 7 || cIdx === 8) && typeof val === 'number') {
                const ref = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
                if (ws[ref]) ws[ref].z = 'R$ #,##0.00';
            }
        }
        return { wch: Math.min(Math.max(maxLen + 3, 10), 45) };
    });
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reservas');

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const fileName = `reservas_pecas_${pad(now.getDate())}-${pad(now.getMonth() + 1)}-${now.getFullYear()}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

document.getElementById('exportExcelBtn')?.addEventListener('click', (e) => {
    e.stopPropagation();
    exportToExcel();
});

/* Filtro rápido ao clicar na tabela */
document.querySelector('#grid tbody')?.addEventListener('click', (e) => {
    const cellVend = e.target.closest('.cell-vendedor');
    if (cellVend) {
        const name = (cellVend.textContent || '').trim();
        if (name && name !== '—' && name !== 'Não informado') {
            const set = new Set(state.Vendedor);
            if (set.has(name)) set.delete(name); else set.add(name);
            state.Vendedor = [...set];
            pillInstances.Vendedor.sync();
            withStableScroll(() => applyFilters());
        }
        return;
    }

    const cellCod = e.target.closest('.cell-codigo');
    if (cellCod) {
        const cod = (cellCod.textContent || '').trim();
        if (cod && cod !== '—') {
            const set = new Set(state.Codigo);
            if (set.has(cod)) set.delete(cod); else set.add(cod);
            state.Codigo = [...set];
            pillInstances.Codigo.sync();
            withStableScroll(() => applyFilters());
        }
        return;
    }
});

/* ===== Inicialização com Dados Pré-Carregados ===== */
// Campos dos registros: ver COLUMN_MAP (leitura pelo cabeçalho da linha 3)
(function init() {
    if (typeof INITIAL_DATA !== 'undefined' && Array.isArray(INITIAL_DATA) && INITIAL_DATA.length > 0) {
        allData = INITIAL_DATA.map(r => ({
            ...r,
            Consultor:    r.Consultor || r.Vendedor || 'Não informado',
            Vendedor:     r.Vendedor || r.Consultor || 'Não informado',
            MesAno:       monthYearFromCell(r.DataReserva),
            FaixaDias:    mapFaixaDias(r.Dias),
            LiberacaoRaw: r.Liberacao,
            Liberacao:    normalizeLiberacao(r.Liberacao),
            Observacao:   r.Observacao || '—'
        }));
        populateFilterOptions();
        applyFilters();
    }
})();
