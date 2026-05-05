// ================================================
// CONFIG
// ================================================
const CAMINHO_ARQUIVO = 'Data/Relatorio_base_EF/Radar.xlsx';

let dadosProcessados = null;
let unidadeAtual = 'SP';

const TRIBUTACOES = [
    'Lucro Presumido',
    'Lucro Real',
    'Simples Nacional'
];

const MAPA_UNIDADE = {
    "SP": "SP",
    "RJ": "RJ",
    "Santos": "Santos",
    "Goias": "GOIAS"
};

// ================================================
// VARIÁVEIS GLOBAIS DO MODAL
// ================================================
let currentModalData = [];
let currentFilterType = 'all';
let currentFilterDate = null;
let currentContexto = '';

// ================================================
// DATAS
// ================================================
function getDiasUteisAbril2026() {
    const dias = [];
    for (let i = 1; i <= 30; i++) {
        const d = new Date(2026, 3, i);
        if (d.getDay() !== 0 && d.getDay() !== 6) dias.push(d);
    }
    return dias;
}

function formatarData(d) {
    return d.toLocaleDateString('pt-BR');
}

function isMesmaData(a, b) {
    return a && b &&
        a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}

function extrairData(valor) {
    if (!valor) return null;

    try {
        if (typeof valor === 'number') {
            const d = XLSX.SSF.parse_date_code(valor);
            return new Date(d.y, d.m - 1, d.d);
        }

        if (valor instanceof Date) {
            return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
        }

        if (typeof valor === 'string') {
            const partes = valor.split(' ')[0];
            const [dia, mes, ano] = partes.split('/');

            if (dia && mes && ano) {
                return new Date(+ano, mes - 1, +dia);
            }

            const d = new Date(valor);
            if (!isNaN(d)) return d;
        }
    } catch (e) {
        console.warn("Erro ao converter data:", valor);
    }

    return null;
}

// ================================================
// CARREGAMENTO
// ================================================
async function carregarArquivo() {
    console.log("📥 Tentando carregar arquivo:", CAMINHO_ARQUIVO);

    const response = await fetch(CAMINHO_ARQUIVO);

    if (!response.ok) {
        throw new Error(`Arquivo não encontrado (${response.status})`);
    }

    const buffer = await response.arrayBuffer();

    const wb = XLSX.read(buffer, {
        type: 'array',
        cellDates: true
    });

    const nomeAba = wb.SheetNames[0];
    const sheet = wb.Sheets[nomeAba];

    const dados = XLSX.utils.sheet_to_json(sheet, {
        raw: true
    });

    return dados;
}

// ================================================
// NORMALIZAÇÃO
// ================================================
function normalizar(dados) {
    return dados.map(r => {
        const n = {};

        Object.keys(r).forEach(k => {
            const key = k.toLowerCase();

            if (key.includes('regimetributario')) n.Tributacao = r[k];
            else if (key.includes('databaixa')) n.DataBaixa = r[k];
            else if (key.includes('datadocumentacao')) n.DataDocumentacao = r[k];
            else if (key.includes('statusdocumentacao')) n.Documentacao = r[k];
            else if (key.includes('documentacaopendente')) n.DocumentacaoPendente = r[k];
            else if (key.includes('unidade')) n.Unidade = r[k];
            else if (key.includes('departamento')) n.Departamento = r[k];
            else if (key.includes('titulo')) n.Titulo = r[k];
            else if (key.includes('usuarioresponsavel')) n.UsuarioResponsavel = r[k];
            else if (key.includes('grupo')) n.Grupo = r[k];
            else if (key.includes('codcliente')) n.CodCliente = r[k];
            else if (key.includes('razaosocial')) n.Cliente = r[k];
        });

        return n;
    });
}

function normalizarDepartamento(dep) {
    if (!dep) return 'Outros';
    return dep;
}

function deveUsarSegmento(unidade) {
    return unidade === 'SP' || unidade === 'Goias';
}

// ================================================
// FILTRO
// ================================================
function filtrarPorUnidade(dados, unidade) {
    const valor = MAPA_UNIDADE[unidade];
    return dados.filter(r => r.Unidade === valor);
}

// ================================================
// CÁLCULO
// ================================================
function calcularEvolucao(base, dias, campo) {
    const total = base.length;
    let pendente = total;
    const arr = [total];

    dias.forEach(dia => {
        const baixados = base.filter(r => {
            const d = extrairData(r[campo]);
            return d && isMesmaData(d, dia);
        }).length;
        pendente -= baixados;
        arr.push(pendente);
    });

    return arr;
}

function calcularDocumentacao(base, dias) {
    // Filtra apenas tarefas com Documentação NÃO "Recebida"
    const baseComDocPendente = base.filter(r => {
        const status = String(r.Documentacao || '').trim().toLowerCase();
        return status !== 'recebida';
    });
    
    const total = baseComDocPendente.length;
    let pendente = total;
    const resultado = [total];

    dias.forEach(dia => {
        const baixados = baseComDocPendente.filter(r => {
            const data = extrairData(r.DataDocumentacao);
            return data && isMesmaData(data, dia);
        }).length;
        pendente -= baixados;
        if (pendente < 0) pendente = 0;
        resultado.push(pendente);
    });

    return resultado;
}

function calcularPendenciaOperacaoReal(base, dias) {
    // Filtra apenas tarefas com Documentação "Recebida"
    const baseComDocRecebida = base.filter(r => {
        const status = String(r.Documentacao || '').trim().toLowerCase();
        return status === 'recebida';
    });
    
    const pendImportacao = calcularEvolucao(baseComDocRecebida, dias, 'DataBaixa');
    return pendImportacao;
}

function calcularPercentual(base, dias) {
    const totalEvolucao = calcularEvolucao(base, dias, 'DataBaixa');
    const total = base.length;
    return {
        pend: totalEvolucao.map(v => (v / total) * 100),
        conc: totalEvolucao.map(v => ((total - v) / total) * 100)
    };
}

function agruparDados(dados, unidade) {
    const usarDepartamento = unidade === 'SP' || unidade === 'Goias';

    if (!usarDepartamento) {
        return { 'Todas as Pendências': dados };
    }

    const grupos = {};
    dados.forEach(r => {
        const chave = normalizarDepartamento(r.Departamento);
        if (!grupos[chave]) grupos[chave] = [];
        grupos[chave].push(r);
    });

    return grupos;
}

// ================================================
// FORMATAÇÃO
// ================================================
function formatarPercentual(valor) {
    return `${Math.round(valor)}%`;
}

function gerarTextoVariacao(valorAtual, valorAnterior) {
    if (valorAnterior === undefined || valorAtual === undefined) return '';
    const diferenca = valorAnterior - valorAtual;
    if (diferenca > 0) {
        return `<span class="variation">▼${diferenca}</span> `;
    }
    return '';
}

// ================================================
// MODAL - FUNÇÕES
// ================================================
function renderModalTable(dataList) {
    const tbody = document.getElementById('modalTableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    if (dataList.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="8" style="text-align: center; padding: 40px;">Nenhuma pendência encontrada</td>';
        tbody.appendChild(tr);
        return;
    }
    
    dataList.forEach((r, index) => {
        const tr = document.createElement('tr');
        tr.className = 'modal-table-row';
        tr.innerHTML = `
            <td>${r.CodCliente || '-'}</td>
            <td>${r.Cliente || '-'}</td>
            <td>${r.Unidade || '-'}</td>
            <td>${r.Tributacao || '-'}</td>
            <td>${r.Titulo || '-'}</td>
            <td>${r.UsuarioResponsavel || '-'}</td>
            <td>${r.Grupo || '-'}</td>
            <td>${r.Documentacao || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function abrirModal(base, data, contexto = '', tipoFiltro = 'all') {
    console.log("🔍 abrirModal chamada!", { baseLength: base.length, data, contexto, tipoFiltro });
    
    let lista = [];
    
    // Primeiro aplica o filtro de data (se houver)
    let baseFiltrada = base;
    if (data) {
        baseFiltrada = base.filter(r => {
            const DataBaixa = extrairData(r.DataBaixa);
            return !DataBaixa || DataBaixa > data;
        });
    }
    
    if (tipoFiltro === 'doc') {
        // Pendência de DOC: Documentação NÃO é "Recebida"
        lista = baseFiltrada.filter(r => {
            const status = String(r.Documentacao || '').trim().toLowerCase();
            return status !== 'recebida';
        });
    } 
    else if (tipoFiltro === 'op') {
        // Pendência de OP: Documentação É "Recebida"
        lista = baseFiltrada.filter(r => {
            const status = String(r.Documentacao || '').trim().toLowerCase();
            return status === 'recebida';
        });
    }
    else {
        // Filtro padrão (todas pendências)
        lista = baseFiltrada;
    }
    
    console.log(`📋 Lista filtrada (${tipoFiltro}):`, lista.length);
    
    // Armazena dados globais para pesquisa e exportação
    currentModalData = lista;
    currentFilterType = tipoFiltro;
    currentFilterDate = data;
    currentContexto = contexto;
    
    // Atualiza título do modal baseado no filtro
    const modalTitle = document.querySelector('#modal .modal-header h2');
    if (modalTitle) {
        let titleText = '📋 Detalhe das Pendências';
        if (tipoFiltro === 'doc') titleText = 'Pendências de Documentação';
        if (tipoFiltro === 'op') titleText = 'Pendências de Operação';
        modalTitle.textContent = titleText;
    }
    
    // Atualiza a tabela
    renderModalTable(lista);
    
    // Limpa e reseta a pesquisa
    const searchInput = document.getElementById('modalSearchInput');
    if (searchInput) searchInput.value = '';
    
    const modal = document.getElementById('modal');
    if (modal) {
        modal.style.display = 'block';
        console.log("✅ Modal aberto com sucesso!");
    }
}

function setupModalSearch() {
    const searchInput = document.getElementById('modalSearchInput');
    if (!searchInput) return;
    
    searchInput.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase().trim();
        const rows = document.querySelectorAll('#modalTableBody .modal-table-row');
        
        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            if (searchTerm === '' || text.includes(searchTerm)) {
                row.classList.remove('hidden');
            } else {
                row.classList.add('hidden');
            }
        });
        
        const visibleRows = document.querySelectorAll('#modalTableBody .modal-table-row:not(.hidden)');
        const noResultsMsg = document.getElementById('noResultsMsg');
        
        if (visibleRows.length === 0 && searchTerm !== '') {
            if (!noResultsMsg) {
                const tbody = document.getElementById('modalTableBody');
                const msgRow = document.createElement('tr');
                msgRow.id = 'noResultsMsg';
                msgRow.innerHTML = '<td colspan="7" style="text-align: center; padding: 40px;">🔍 Nenhum resultado encontrado para "' + searchTerm + '"</td>';
                tbody.appendChild(msgRow);
            }
        } else {
            const msgRow = document.getElementById('noResultsMsg');
            if (msgRow) msgRow.remove();
        }
    });
}

function exportToExcel() {
    if (!currentModalData || currentModalData.length === 0) {
        alert('Nenhum dado para exportar!');
        return;
    }
    
    const exportData = currentModalData.map(r => ({
        'ID Cliente': r.CodCliente || '-',
        'Cliente': r.Cliente || '-',
        'Grupo': r.Grupo || '-',
        'Gerente': r.Gerente || '-',
        'Tributação': r.Tributacao || '-',
        'Equipe': r.EquipeAtendimento || '-',
        'Segmento': r.Segmento || '-',
        'Data Importação': r.DataBaixa ? formatarData(extrairData(r.DataBaixa)) : '-',
        'Documentação': r.Documentacao || '-',
        'Status DOC': r.Documentacao === 'Recebida' ? 'Recebida' : 'Pendente'
    }));
    
    const ws = XLSX.utils.json_to_sheet(exportData);
    const colWidths = [
        {wch:12}, {wch:30}, {wch:15}, {wch:15}, {wch:15}, {wch:15}, {wch:12}, {wch:15}, {wch:12}, {wch:12}
    ];
    ws['!cols'] = colWidths;
    
    const wb = XLSX.utils.book_new();
    const sheetName = `Pendencias_${currentFilterType}_${new Date().toISOString().slice(0,19)}`;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    XLSX.writeFile(wb, `pendencias_${currentFilterType}_${currentFilterDate ? formatarData(currentFilterDate) : 'inicio'}.xlsx`);
}

// ================================================
// CRIAÇÃO DE LINHAS
// ================================================
function criarLinha(nome, valores, base, dias, isPercentual = false, isTotalOuTributacao = false, tipoModal = 'all') {
    const tr = document.createElement('tr');

    const tdNome = document.createElement('td');
    tdNome.textContent = nome;
    tr.appendChild(tdNome);

    valores.forEach((v, i) => {
        const td = document.createElement('td');
        td.className = 'clickable';

        // 🔎 Define data da coluna
        const dataColuna = i === 0 ? null : dias[i - 1];

        // 🧠 Verifica se é data futura
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);

        let futura = false;
        if (dataColuna) {
            const d = new Date(dataColuna);
            d.setHours(0, 0, 0, 0);
            futura = d > hoje;
        }

        // 🎯 Define valor a exibir
        let valorDisplay = '';

        if (!futura) {
            if (isPercentual) {
                valorDisplay = formatarPercentual(v);
            } else {
                valorDisplay = Math.round(v);
            }
        }

        // 🔢 Variação (apenas se não for futura)
        if (!futura && isTotalOuTributacao && i > 0 && !isPercentual) {
            const valorAnterior = valores[i - 1];
            const textoVariacao = gerarTextoVariacao(v, valorAnterior);

            if (textoVariacao) {
                td.innerHTML = `${textoVariacao}${valorDisplay}`;
            } else {
                td.textContent = valorDisplay;
            }
        } else {
            td.textContent = valorDisplay;
        }



        // 📊 Define base para modal
        let baseParaModal = base;
        if (tipoModal === 'doc') {
            baseParaModal = base.filter(r => {
                const status = String(r.Documentacao || '').trim().toLowerCase();
                return status !== 'recebida';
            });
        } else if (tipoModal === 'op') {
            baseParaModal = base.filter(r => {
                const status = String(r.Documentacao || '').trim().toLowerCase();
                return status === 'recebida';
            });
        }

        // 🖱️ Evento de clique (somente se NÃO for futura)
        if (!futura) {
            td.onclick = (e) => {
                e.stopPropagation();
                const dataSelecionada = i === 0 ? null : dias[i - 1];
                window.abrirModal(baseParaModal, dataSelecionada, nome, tipoModal);
            };
        }

        tr.appendChild(td);
    });

    return tr;
}

function criarLinhaPercentual(nome, valores, tipo, dias = []) {
    const tr = document.createElement('tr');
    
    const tdNome = document.createElement('td');
    tdNome.textContent = nome;
    tr.appendChild(tdNome);
    
    valores.forEach((v, i) => {
        const td = document.createElement('td');
        td.style.padding = '0';
        td.style.position = 'relative';

        // 🔎 Segurança total aqui
        const dataColuna = (i === 0 || !dias || !dias[i - 1]) ? null : dias[i - 1];

        // 🧠 Verifica se é futura
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);

        let futura = false;
        if (dataColuna) {
            const d = new Date(dataColuna);
            d.setHours(0, 0, 0, 0);
            futura = d > hoje;
        }

        // 🚫 Se for futura → célula vazia
        if (futura) {
            td.style.backgroundColor = '#f5f5f5';
            td.style.opacity = '0.6';
            tr.appendChild(td);
            return;
        }

        const wrapper = document.createElement('div');
        wrapper.className = 'percent-wrapper';
        wrapper.style.position = 'relative';
        wrapper.style.display = 'flex';
        wrapper.style.alignItems = 'center';
        wrapper.style.justifyContent = 'flex-end';
        wrapper.style.minHeight = '40px';
        
        const barBg = document.createElement('div');
        barBg.className = `percent-bar-bg ${tipo === 'danger' ? 'bar-danger' : 'bar-success'}`;
        const percentValue = Math.min(Math.round(v), 100);
        barBg.style.width = `${percentValue}%`;
        barBg.style.position = 'absolute';
        barBg.style.top = '50%';
        barBg.style.transform = 'translateY(-50%)';
        barBg.style.left = '0';
        barBg.style.height = '35px';
        barBg.style.borderRadius = '3px';
        
        const valueSpan = document.createElement('span');
        valueSpan.className = 'percent-value';
        valueSpan.style.position = 'relative';
        valueSpan.style.zIndex = '2';
        valueSpan.style.fontWeight = '700';
        valueSpan.style.fontSize = '0.9em';
        valueSpan.style.color = '#000000';
        valueSpan.textContent = formatarPercentual(v);
        
        wrapper.appendChild(barBg);
        wrapper.appendChild(valueSpan);
        td.appendChild(wrapper);
        tr.appendChild(td);
    });
    
    return tr;
}

function ehDataFutura(data) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const d = new Date(data);
    d.setHours(0, 0, 0, 0);

    return d > hoje;
}

// ================================================
// CRIAÇÃO DO BLOCO/QUADRO
// ================================================
function criarBloco(nomeGrupo, dados, dias) {
    const container = document.createElement('div');
    container.className = 'tributacao-section';

    const header = document.createElement('div');
    header.className = 'section-header';
    const titulo = document.createElement('h2');
    titulo.textContent = nomeGrupo;
    header.appendChild(titulo);
    container.appendChild(header);

    const scrollWrapper = document.createElement('div');
    scrollWrapper.className = 'scroll-wrapper';

    const tabela = document.createElement('table');
    tabela.className = 'dashboard-table';

    const thead = document.createElement('thead');
    const trHead = document.createElement('tr');
    trHead.innerHTML = '<th>Indicador</th><th>Início</th>';

    dias.forEach(d => {
        const th = document.createElement('th');
        th.textContent = formatarData(d);
        trHead.appendChild(th);
    });

    thead.appendChild(trHead);
    tabela.appendChild(thead);

    const tbody = document.createElement('tbody');

    // TOTAL
    const totalLinha = criarLinha('+ Total', calcularEvolucao(dados, dias, 'DataBaixa'), dados, dias, false, true, 'all');
    totalLinha.classList.add('linha-total');
    tbody.appendChild(totalLinha);

    // PERCENTUAIS
    const perc = calcularPercentual(dados, dias);
    const linhaPercPend = criarLinhaPercentual('% Pendente', perc.pend, 'danger', dias);
    const linhaPercConc = criarLinhaPercentual('% Concluída', perc.conc, 'success', dias);
    tbody.appendChild(linhaPercPend);
    tbody.appendChild(linhaPercConc);

    // =========================
    // DRILL DOWN
    // =========================
    let expandidoTotal = false;

    totalLinha.onclick = (e) => {
        e.stopPropagation();

        expandidoTotal = !expandidoTotal;
        totalLinha.children[0].textContent = expandidoTotal ? '- Total' : '+ Total';

        // remove tudo abaixo
        tbody.querySelectorAll('.nivel1, .nivel2, .nivel3, .separator-row').forEach(e => e.remove());

        if (!expandidoTotal) return;

        TRIBUTACOES.forEach(trib => {
            const base = dados.filter(d => d.Tributacao === trib);
            if (base.length === 0) return;

            const linhaTrib = criarLinha(`+ ${trib}`, calcularEvolucao(base, dias, 'DataBaixa'), base, dias, false, true, 'all');
            linhaTrib.classList.add('linha-tributacao', 'nivel1');

            let expandidoTrib = false;

            linhaTrib.onclick = (e) => {
                e.stopPropagation();

                expandidoTrib = !expandidoTrib;
                linhaTrib.children[0].textContent = expandidoTrib ? `- ${trib}` : `+ ${trib}`;

                // remove nível abaixo
                let next = linhaTrib.nextSibling;
                while (next && (next.classList.contains('nivel2') || next.classList.contains('nivel3') || next.classList.contains('separator-row'))) {
                    const temp = next;
                    next = next.nextSibling;
                    temp.remove();
                }

                if (!expandidoTrib) return;

                // =========================
                // AGRUPAR POR TAREFA
                // =========================
                const tarefas = {};

                base.forEach(r => {
                    const chave = r.Titulo || 'Outros';
                    if (!tarefas[chave]) tarefas[chave] = [];
                    tarefas[chave].push(r);
                });

                let referencia = linhaTrib;

                Object.keys(tarefas).forEach(tarefa => {
                    const baseTarefa = tarefas[tarefa];

                    const linhaTarefa = criarLinha(`+ ${tarefa}`, calcularEvolucao(baseTarefa, dias, 'DataBaixa'), baseTarefa, dias);
                    linhaTarefa.classList.add('nivel2');
                    linhaTarefa.children[0].style.paddingLeft = '30px';

                    let expandidoTarefa = false;

                    linhaTarefa.onclick = (e) => {
                        e.stopPropagation();

                        expandidoTarefa = !expandidoTarefa;
                        linhaTarefa.children[0].textContent = expandidoTarefa ? `- ${tarefa}` : `+ ${tarefa}`;

                        // remove subnível
                        let next = linhaTarefa.nextSibling;
                        while (next && next.classList.contains('nivel3')) {
                            const temp = next;
                            next = next.nextSibling;
                            temp.remove();
                        }

                        if (!expandidoTarefa) return;

                        const doc = calcularDocumentacao(baseTarefa, dias);
                        const op = calcularPendenciaOperacaoReal(baseTarefa, dias);
                        const perc = calcularPercentual(baseTarefa, dias);

                        const subLinhas = [
                            criarLinha('Doc Pendente', doc, baseTarefa, dias, false, false, 'doc'),
                            criarLinha('Pendência OP', op, baseTarefa, dias, false, false, 'op'),
                            criarLinhaPercentual('% Pendente', perc.pend, 'danger', dias),
                            criarLinhaPercentual('% Concluído', perc.conc, 'success', dias)
                        ];

                        let refInterno = linhaTarefa;

                        subLinhas.forEach(l => {
                            l.classList.add('nivel3');
                            l.children[0].style.paddingLeft = '50px';
                            refInterno.parentNode.insertBefore(l, refInterno.nextSibling);
                            refInterno = l;
                        });
                    };

                    referencia.parentNode.insertBefore(linhaTarefa, referencia.nextSibling);
                    referencia = linhaTarefa;
                });

                // linha separadora
                const separatorRow = document.createElement('tr');
                separatorRow.classList.add('separator-row', 'nivel2');

                const tdSeparator = document.createElement('td');
                tdSeparator.setAttribute('colspan', dias.length + 2);
                tdSeparator.style.height = '12px';
                tdSeparator.style.backgroundColor = '#f4f6fa';
                tdSeparator.style.border = 'none';

                separatorRow.appendChild(tdSeparator);
                referencia.parentNode.insertBefore(separatorRow, referencia.nextSibling);
            };

            tbody.insertBefore(linhaTrib, linhaPercPend);
        });
    };

    tabela.appendChild(tbody);
    scrollWrapper.appendChild(tabela);
    container.appendChild(scrollWrapper);

    return container;
}

// ================================================
// ATUALIZAR DASHBOARD
// ================================================
function atualizarDashboard(dados, dias, unidade) {
    const container = document.getElementById('dashboards-container');
    if (!container) return;
    container.innerHTML = '';
    const filtrados = filtrarPorUnidade(dados, unidade);
    const grupos = agruparDados(filtrados, unidade);
    Object.keys(grupos).forEach(nome => {
        container.appendChild(criarBloco(nome, grupos[nome], dias));
    });
}

// ================================================
// INIT
// ================================================
document.addEventListener('DOMContentLoaded', async () => {
    const dias = getDiasUteisAbril2026();
    const status = document.getElementById('statusMessage');

    try {
        let dados = await carregarArquivo();
        dados = normalizar(dados);
        dadosProcessados = dados;
        atualizarDashboard(dados, dias, 'SP');
        if (status) {
            status.innerHTML = '✅ Arquivo carregado com sucesso';
            status.style.color = '#27ae60';
        }
    } catch (e) {
        if (status) {
            status.innerHTML = '❌ Erro ao carregar arquivo';
            status.style.color = '#dc3545';
        }
        console.error(e);
    }

    // FILTRO POR BOTÕES
    const botoes = document.querySelectorAll('.btn-unidade');
    botoes.forEach(btn => {
        btn.addEventListener('click', () => {
            botoes.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const unidade = btn.getAttribute('data-unidade');
            unidadeAtual = unidade;
            if (dadosProcessados) {
                atualizarDashboard(dadosProcessados, dias, unidade);
            }
        });
    });

    // Configurar pesquisa no modal
    setupModalSearch();

    // Configurar botão de exportar Excel
    const exportBtn = document.getElementById('exportExcelBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportToExcel);
    }

    // FECHAR MODAL
    const closeModalBtn = document.querySelector('.close-modal');
    const modal = document.getElementById('modal');
    
    if (closeModalBtn) {
        closeModalBtn.onclick = () => {
            if (modal) modal.style.display = 'none';
        };
    }
    
    window.onclick = (e) => {
        if (e.target === modal) {
            if (modal) modal.style.display = 'none';
        }
    };
    
    console.log("✅ Dashboard inicializado");
});