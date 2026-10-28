
let charts = [];
let paginaAnterior = "dashboard";
let paginaAtual = "dashboard";
let dadosOriginais = [];
let selecionado = null;
let itemSelecionado = null;
let estoque = []
let equipamentos = [];
let equipamentoSelecionado = null;
let modoEdicaoEquip = false;
let setores = [
    "Corte",
    "Correr",
    "Cravação",
    "Vidro",
    "Expedição",
    "Logística",
    "Bites",
    "Manutenção",
    "Outros"
];

let termoProcuraAtivo = "";
let ordenacaoAtual = { coluna: null, direcao: "asc" };

const CHAVE_ULTIMO_EXCEL = "ultimo_excel_base64";
const CHAVE_DADOS_LOCAL = "dashboard_manutencao_dados";

document.getElementById('tipoVisualizacao').addEventListener('change', atualizarFiltrosDashboard);
document.getElementById('anoFiltro').addEventListener('change', atualizarFiltrosDashboard);
document.getElementById('mesFiltro').addEventListener('change', atualizarFiltrosDashboard);
document.getElementById('diaFiltro').addEventListener('change', aplicarFiltrosDashboard);
document.getElementById('percentualMeta').addEventListener('input', aplicarFiltrosDashboard);

window.addEventListener("load", async () => {
    const dadosLocaisCarregados = carregarDadosGuardados();
    carregarEstoque();
    carregarEquipamentos();

    // Não sobrescrever dados locais válidos com uma cópia antiga do Excel.
    if (!dadosLocaisCarregados) {
        if (localStorage.getItem(CHAVE_ULTIMO_EXCEL)) {
            await carregarUltimoExcel();
        } else {
            await carregarExcelOnline();
        }
    }

    preencherFiltroAno(dadosOriginais);
    atualizarFiltrosDashboard();
    atualizarListaManutencao();
    atualizarHistoricoManutencao();
    atualizarTabelaEstoque();
    atualizarTabelaEquipamentos();
    atualizarDropdownEquipamentos();
    atualizarDropdownSetores();
});

document.addEventListener("change", function (e) {
    if (e.target?.id !== "novoEquipamento") return;
    const eq = equipamentos.find(x => x.nome === e.target.value);
    const setor = document.getElementById("novoSetor");
    if (eq && setor) setor.value = eq.setor;
});


/* NAVEGAÇÃO */
function mostrarPagina(pagina) {
    paginaAtual = pagina;
    
    let painel = document.getElementById("painelManutencao");
    if (painel) painel.style.display = "none";
    
    let atual = document.querySelector(".page.active");
    if (atual) {
        paginaAnterior = atual.id;
    }

    document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
    document.getElementById(pagina).classList.add("active");

    if (pagina === "manutencao") {
        atualizarListaManutencao();
    }

    if (pagina === "historico") {
        atualizarHistoricoManutencao();
    }

    if (pagina === "estoque") {
        atualizarTabelaEstoque();
    }

    if (pagina === "configuracoes") {
        atualizarTabelaEquipamentos();
    }
}
/* NORMALIZAR TEXTO */
function normalizarTexto(txt) {
    return String(txt || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim();
}
/* DETECTAR COLUNA */
function detectarColuna(row, nomes) {
    for (let key in row) {
        let keyNormalizada = normalizarTexto(key);

        for (let nome of nomes) {
            if (keyNormalizada.includes(normalizarTexto(nome))) {
                return row[key];
            }
        }
    }

    return null;
}
/* GRAVAR CAMPO SEM QUEBRAR NOMES ORIGINAIS */
function definirCampo(row, nomes, nomePadrao, valor) {
    for (let key in row) {
        let keyNormalizada = normalizarTexto(key);

        for (let nome of nomes) {
            if (keyNormalizada.includes(normalizarTexto(nome))) {
                row[key] = valor;
                return;
            }
        }
    }

    row[nomePadrao] = valor;
}
/* PEGAR CAMPOS */
function getData(row) {

    // ✅ PRIORIDADE 1 — campo já certo
    if (row["Data de Inserção"]) return row["Data de Inserção"];

    // ✅ PRIORIDADE 2 — hora início (TEU CASO REAL)
    if (row["Hora de Início"]) return row["Hora de Início"];

    // ✅ fallback automático
    let valor = detectarColuna(row, [
        "data de insercao",
        "data de inserção",
        "data",
        "hora de inicio",
        "hora de início"
    ]);

    return valor || null;
}

function getEquipamento(row) {
    return detectarColuna(row, ["equipamento", "maquina", "máquina"]);
}

function getManutencao(row) {
    return row?.["Manutenção"] ?? row?.["Manutencao"] ?? detectarColuna(row, ["manutencao"]);
}

function getSetor(row) {
    return detectarColuna(row, ["setor", "area", "área"]);
}

function getTipo(row) {
    if (!row) return "";
    const chaves = Object.keys(row);
    const exata = chaves.find(k => ["tipo de manutencao", "tipo"].includes(normalizarTexto(k)));
    if (exata !== undefined) return row[exata];
    const especifica = chaves.find(k => normalizarTexto(k).includes("tipo de manutencao"));
    return especifica !== undefined ? row[especifica] : "";
}

function getFalha(row) {
    return row["Tipo de Falha"] || row["Falha"] || row.falha || "";
}

function getEstado(row) {

    // ✅ prioridade total
    if (row["Estado"]) return row["Estado"];

    // ✅ fallback
    let valor = detectarColuna(row, ["estado"]);

    return valor || "";
}

function getDescricao(row) {
    return detectarColuna(row, ["descricao", "descrição"]);
}

function getDataConclusao(row) {

    if (row["Data de Conclusão"]) return row["Data de Conclusão"];

    if (row["Hora de Conclusão"]) return row["Hora de Conclusão"];

    let valor = detectarColuna(row, [
        "data de conclusao",
        "data de conclusão",
        "conclusao",
        "conclusão",
        "hora de conclusao",
        "hora de conclusão"
    ]);

    return valor || null;
}

function getHoraInicio(row) {
    return detectarColuna(row, [
        "hora de inicio",
        "hora de início",
        "hora inicio",
        "hora início",
        "inicio",
        "início",
        "inicio defeito",
        "início defeito"
    ]);
}

function getHoraConclusao(row) {
    return detectarColuna(row, [
        "hora de conclusao",
        "hora de conclusão",
        "hora conclusao",
        "hora conclusão",
        "fim",
        "fim defeito",
        "hora final"
    ]);
}

function getRecorrente(row) {
    return detectarColuna(row, ["recorrente"]);
}

function getFrequencia(row) {
    return detectarColuna(row, ["frequencia", "frequência"]);
}
/* DATAS */
function excelDateToJSDate(valor) {
    if (valor === null || valor === undefined || valor === "") return null;
    if (valor instanceof Date) return Number.isNaN(valor.getTime()) ? null : new Date(valor.getTime());

    if (typeof valor === "number" && Number.isFinite(valor)) {
        // Serial Excel: preservar a hora local representada pelo ficheiro.
        const dias = Math.floor(valor);
        const fracao = valor - dias;
        const base = new Date(Date.UTC(1899, 11, 30) + dias * 86400000);
        const minutos = Math.round(fracao * 1440);
        const d = new Date(base.getUTCFullYear(), base.getUTCMonth(), base.getUTCDate(), 0, minutos);
        return Number.isNaN(d.getTime()) ? null : d;
    }

    if (typeof valor === "string") {
        const texto = valor.trim();
        if (!texto) return null;
        // ISO (inclui yyyy-mm-dd e yyyy-mm-ddThh:mm)
        const iso = texto.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
        if (iso) {
            const [, ys, ms, ds, hs = "0", mins = "0", secs = "0"] = iso;
            const d = new Date(Number(ys), Number(ms)-1, Number(ds), Number(hs), Number(mins), Number(secs));
            if (d.getFullYear() !== Number(ys) || d.getMonth() !== Number(ms)-1 || d.getDate() !== Number(ds)) return null;
            return d;
        }
        // Português: dd/mm/yyyy [hh:mm[:ss]]
        const pt = texto.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
        if (pt) {
            const [, ds, ms, ys, hs = "0", mins = "0", secs = "0"] = pt;
            const dia=Number(ds), mes=Number(ms), ano=Number(ys), hora=Number(hs), minuto=Number(mins), segundo=Number(secs);
            if (mes < 1 || mes > 12 || dia < 1 || dia > 31 || hora > 23 || minuto > 59 || segundo > 59) return null;
            const d = new Date(ano, mes-1, dia, hora, minuto, segundo);
            if (d.getFullYear() !== ano || d.getMonth() !== mes-1 || d.getDate() !== dia) return null;
            return d;
        }
        // Último recurso para strings ISO com timezone explícito.
        const parsed = new Date(texto);
        return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
}
function formatarData(valor) {
    let d = excelDateToJSDate(valor);
    return d ? d.toLocaleDateString("pt-PT") : "";
}

function formatarDataHora(valor) {
    let d = excelDateToJSDate(valor);
    if (!d) return "";

    let data = d.toLocaleDateString("pt-PT");
    let hora = String(d.getHours()).padStart(2, "0");
    let minuto = String(d.getMinutes()).padStart(2, "0");

    return `${data} ${hora}:${minuto}`;
}

function formatarDataInput(valor) {
    let d = excelDateToJSDate(valor);
    if (!d) return "";

    let ano = d.getFullYear();
    let mes = String(d.getMonth() + 1).padStart(2, "0");
    let dia = String(d.getDate()).padStart(2, "0");

    return `${ano}-${mes}-${dia}`;
}

function formatarDataHoraInput(valor) {
    let d = excelDateToJSDate(valor);
    if (!d) return "";

    let ano = d.getFullYear();
    let mes = String(d.getMonth() + 1).padStart(2, "0");
    let dia = String(d.getDate()).padStart(2, "0");
    let hora = String(d.getHours()).padStart(2, "0");
    let minuto = String(d.getMinutes()).padStart(2, "0");

    return `${ano}-${mes}-${dia}T${hora}:${minuto}`;
}

function formatarHora(valor) {
    let d = excelDateToJSDate(valor);
    if (!d) return "";

    let hora = String(d.getHours()).padStart(2, "0");
    let minuto = String(d.getMinutes()).padStart(2, "0");

    return `${hora}:${minuto}`;
}

function nomeDiaSemana(data) {
    let nomes = ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"];
    return nomes[data.getDay()];
}

function estaConcluida(row) {
    return normalizarTexto(getEstado(row)).includes("concl");
}

function diferencaDias(dataInicial, dataFinal) {
    let inicio = excelDateToJSDate(dataInicial);
    let fim = dataFinal || new Date();

    if (!inicio) return 0;

    let inicioSemHora = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());
    let fimSemHora = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate());

    return Math.floor((fimSemHora - inicioSemHora) / (1000 * 60 * 60 * 24));
}

function adicionarDias(dataBase, dias) {
    let data = excelDateToJSDate(dataBase) || new Date();
    let novaData = new Date(data);
    novaData.setDate(novaData.getDate() + dias);
    return novaData;
}

function adicionarHoras(dataBase, horas) {
    let data = excelDateToJSDate(dataBase) || new Date();
    let novaData = new Date(data);
    novaData.setHours(novaData.getHours() + horas);
    return novaData;
}

function calcularMinutosEntre(inicio, conclusao) {
    let dataInicio = excelDateToJSDate(inicio);
    let dataConclusao = excelDateToJSDate(conclusao);

    if (!dataInicio || !dataConclusao) return null;

    let diferenca = dataConclusao - dataInicio;

    if (diferenca < 0) return null;

    return Math.round(diferenca / 60000);
}

function formatarDuracao(minutos) {
    if (minutos === null || minutos === undefined || isNaN(minutos)) return "0h";

    let horas = Math.floor(minutos / 60);
    let mins = minutos % 60;

    if (horas > 0 && mins > 0) return `${horas}h ${mins}m`;
    if (horas > 0) return `${horas}h`;
    return `${mins}m`;
}
/* NORMALIZAR DADOS ANTIGOS */
function normalizarHorariosExistentes() {
    // Normalização conservadora: não inventar horários para registos históricos.
    dadosOriginais.forEach(row => {
        if (!row || typeof row !== "object") return;
        const inicio = getHoraInicio(row) || getData(row);
        if (inicio) {
            const dataInicio = excelDateToJSDate(inicio);
            if (dataInicio && !getHoraInicio(row)) definirCampo(row, ["hora de inicio", "hora de início", "inicio", "início"], "Hora de Início", dataInicio);
        }
        const conclusao = getHoraConclusao(row) || getDataConclusao(row);
        if (conclusao) {
            const dataConclusao = excelDateToJSDate(conclusao);
            if (dataConclusao && !getHoraConclusao(row)) definirCampo(row, ["hora de conclusao", "hora de conclusão", "fim"], "Hora de Conclusão", dataConclusao);
        }
    });
}
/* GUARDAR/CARREGAR DADOS LOCALMENTE */
function guardarDadosLocalmente() {
    try {
        localStorage.setItem(CHAVE_DADOS_LOCAL, JSON.stringify(dadosOriginais));
        return true;
    } catch (erro) {
        console.error("Não foi possível guardar as manutenções localmente:", erro);
        mostrarMensagem("Erro ao guardar dados: armazenamento cheio ou indisponível.");
        return false;
    }
}

function carregarDadosGuardados() {
    const dados = localStorage.getItem(CHAVE_DADOS_LOCAL);
    if (dados === null) return false;
    try {
        const parsed = JSON.parse(dados);
        if (!Array.isArray(parsed)) throw new Error("Formato de dados inválido");
        dadosOriginais = parsed;
        normalizarHorariosExistentes();
        preencherFiltroAno(dadosOriginais);
        atualizarFiltrosDashboard();
        atualizarListaManutencao();
        atualizarHistoricoManutencao();
        return true;
    } catch (erro) {
        console.error("Erro ao carregar dados guardados:", erro);
        mostrarMensagem("Os dados locais estão inválidos. Importe um Excel válido.");
        return false;
    }
}
/* DESTRUIR GRÁFICOS */
function destroyCharts() {
    charts.forEach(c => c.destroy());
    charts = [];
}
/* FILTROS DO DASHBOARD */
function preencherFiltroAno(data) {
    let anos = new Set();

    data.forEach(row => {
        let d = excelDateToJSDate(getData(row));
        if (d) anos.add(d.getFullYear());
    });

    let select = document.getElementById("anoFiltro");
    let valorAnterior = select.value;
    select.innerHTML = "";

    let anosArray = Array.from(anos).sort((a,b) => a-b);

    anosArray.forEach(ano => {
        let option = document.createElement("option");
        option.value = ano;
        option.textContent = ano;
        select.appendChild(option);
    });

    if (anosArray.length > 0) {
        if (valorAnterior && anos.has(Number(valorAnterior))) {
            select.value = valorAnterior;
        } else {
            select.value = Math.max(...anosArray);
        }
    }
}

function atualizarFiltrosDashboard() {
    preencherFiltroMes();
    preencherFiltroDia();
    controlarFiltrosVisuais();
    aplicarFiltrosDashboard();
}

function preencherFiltroMes() {
    let ano = document.getElementById("anoFiltro").value;
    let selectMes = document.getElementById("mesFiltro");
    let mesAtualSelecionado = selectMes.value;

    let meses = new Set();
    let nomesMeses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

    dadosOriginais.forEach(row => {
        let d = excelDateToJSDate(getData(row));
        if (!d) return;
        if (ano && d.getFullYear() != ano) return;

        meses.add(d.getMonth());
    });

    selectMes.innerHTML = "";

    let mesesArray = Array.from(meses).sort((a,b) => a-b);

    mesesArray.forEach(m => {
        let option = document.createElement("option");
        option.value = m;
        option.textContent = nomesMeses[m];
        selectMes.appendChild(option);
    });

    if (mesesArray.length > 0) {
        if (mesAtualSelecionado !== "" && meses.has(Number(mesAtualSelecionado))) {
            selectMes.value = mesAtualSelecionado;
        } else {
            selectMes.value = mesesArray[0];
        }
    }
}

function preencherFiltroDia() {
    let ano = document.getElementById("anoFiltro").value;
    let mes = document.getElementById("mesFiltro").value;
    let selectDia = document.getElementById("diaFiltro");
    let diaAtualSelecionado = selectDia.value;

    let dias = new Set();

    dadosOriginais.forEach(row => {
        let d = excelDateToJSDate(getData(row));
        if (!d) return;
        if (ano && d.getFullYear() != ano) return;
        if (mes !== "" && d.getMonth() != mes) return;

        dias.add(d.getDate());
    });

    selectDia.innerHTML = "";

    let diasArray = Array.from(dias).sort((a,b) => a-b);

    diasArray.forEach(dia => {
        let option = document.createElement("option");
        option.value = dia;
        option.textContent = String(dia).padStart(2, "0");
        selectDia.appendChild(option);
    });

    if (diasArray.length > 0) {
        if (diaAtualSelecionado !== "" && dias.has(Number(diaAtualSelecionado))) {
            selectDia.value = diaAtualSelecionado;
        } else {
            selectDia.value = diasArray[0];
        }
    }
}

function controlarFiltrosVisuais() {
    let tipo = document.getElementById("tipoVisualizacao").value;
    let mesFiltro = document.getElementById("mesFiltro");
    let diaFiltro = document.getElementById("diaFiltro");

    mesFiltro.disabled = tipo === "ano";
    diaFiltro.disabled = tipo !== "dia";
}

function aplicarFiltrosDashboard() {
    let tipoVisualizacao = document.getElementById("tipoVisualizacao").value;
    let ano = document.getElementById("anoFiltro").value;
    let mes = document.getElementById("mesFiltro").value;
    let dia = document.getElementById("diaFiltro").value;

    let filtrados = dadosOriginais.filter(row => {
        let d = excelDateToJSDate(getData(row));
        if (!d) return false;

        if (ano && d.getFullYear() != ano) return false;

        if ((tipoVisualizacao === "mes" || tipoVisualizacao === "dia") && mes !== "" && d.getMonth() != mes) return false;

        if (tipoVisualizacao === "dia" && dia !== "" && d.getDate() != dia) return false;

        return true;
    });

    processarDashboard(filtrados, tipoVisualizacao);
}

/* DASHBOARD */
function processarDashboard(data, tipoVisualizacao) {
    destroyCharts();

    let tipo = {};
    let falha = {};
    let setor = {};
    let graficoPrincipal = {};
    let concluidas = 0;
    let pendentes = 0;
    let totalMinutosResposta = 0;
    let totalRespostasValidas = 0;

    let nomesMeses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

    if (tipoVisualizacao === "ano") {
        nomesMeses.forEach(m => graficoPrincipal[m] = 0);
    }

    data.forEach(row => {
        let t = getTipo(row) || "Outro";
        let f = getFalha(row) || "Outro";
        let s = getSetor(row) || "Outro";

        let d = excelDateToJSDate(getData(row));

        tipo[t] = (tipo[t] || 0) + 1;
        falha[f] = (falha[f] || 0) + 1;
        setor[s] = (setor[s] || 0) + 1;

        if (d) {
            let chave = "";

            if (tipoVisualizacao === "ano") {
                chave = nomesMeses[d.getMonth()];
            }

            if (tipoVisualizacao === "mes") {
                chave = `${String(d.getDate()).padStart(2, "0")}/${nomeDiaSemana(d)}`;
            }

            if (tipoVisualizacao === "dia") {
                chave = formatarHora(d);
            }

            graficoPrincipal[chave] = (graficoPrincipal[chave] || 0) + 1;
        }

        if (estaConcluida(row)) {
            concluidas++;

            let minutosResposta = calcularMinutosEntre(getHoraInicio(row), getHoraConclusao(row));

            if (minutosResposta !== null) {
                totalMinutosResposta += minutosResposta;
                totalRespostasValidas++;
            }
        } else {
            pendentes++;
        }
    });

    let mediaResposta = totalRespostasValidas > 0
        ? Math.round(totalMinutosResposta / totalRespostasValidas)
        : 0;

    document.getElementById("concluidas").innerText = concluidas;
    document.getElementById("pendentes").innerText = pendentes;
    document.getElementById("tempoResposta").innerText = formatarDuracao(mediaResposta);

    let tituloGrafico = "Manutenções por Mês";

    if (tipoVisualizacao === "mes") {
        tituloGrafico = "Manutenções por Dia";
    }

    if (tipoVisualizacao === "dia") {
        tituloGrafico = "Manutenções por Hora";
    }

    charts.push(criarBarraPrincipal("porDia", graficoPrincipal, tituloGrafico));
    charts.push(criarPizza("tipoIntervencao", tipo, "Tipos de Intervenções"));
    charts.push(criarPizza("tipoFalha", falha, "Tipos de Falhas"));
    charts.push(criarBarra("setor", setor, "Manutenção por Setor"));

    atualizarLoadingMeta(data, tipoVisualizacao, concluidas);
}

/* GRÁFICOS */
function criarPizza(id, dados, titulo) {

    return new Chart(document.getElementById(id), {
        type: "pie",
        data: {
            labels: Object.keys(dados),
            datasets: [{
                data: Object.values(dados),
                backgroundColor: [
                    "#007bff", "#28a745", "#ffc107", "#dc3545",
                    "#6f42c1", "#20c997", "#fd7e14", "#6c757d"
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: titulo
                },

                datalabels: {
                    color: "#fff",
                    font: {
                        weight: "bold",
                        size: 14
                    },
                    formatter: (value, context) => {

                        let total = context.dataset.data.reduce((a, b) => a + b, 0);

                        let percent = (value / total * 100).toFixed(1) + "%";

                        return percent;
                    }
                }
            }
        },
        plugins: [ChartDataLabels]
    });
}

function criarBarra(id, dados, titulo) {
    return new Chart(document.getElementById(id), {
        type: "bar",
        data: {
            labels: Object.keys(dados),
            datasets: [{
                data: Object.values(dados),
                backgroundColor: [
                    "#007bff", "#28a745", "#ffc107", "#dc3545",
                    "#6f42c1", "#20c997", "#fd7e14", "#6c757d"
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                title: { display: true, text: titulo }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1
                    }
                }
            }
        }
    });
}

function criarBarraPrincipal(id, dados, titulo) {
    
let labels = Object.keys(dados);

labels.sort((a, b) => {

    // ✅ para "dia" e "mês"
    let diaA = parseInt(a.split("/")[0]);
    let diaB = parseInt(b.split("/")[0]);

    return diaA - diaB;
});

    return new Chart(document.getElementById(id), {
        type: "bar",
        data: {
            labels: labels,
            datasets: [{
                label: titulo,
                data: labels.map(d => dados[d]),
                backgroundColor: "#1f77b4"
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: titulo },
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1
                    }
                }
            }
        }
    });
}
/* META EM LOADING */
function obterDadosDoMesSelecionado() {
    let ano = document.getElementById("anoFiltro").value;
    let mes = document.getElementById("mesFiltro").value;

    return dadosOriginais.filter(row => {
        let d = excelDateToJSDate(getData(row));
        if (!d) return false;

        if (ano && d.getFullYear() != ano) return false;
        if (mes !== "" && d.getMonth() != mes) return false;

        return true;
    });
}

function atualizarLoadingMeta(dataFiltrada, tipoVisualizacao, concluidasFiltradas) {
    let percentual = Math.min(100, Math.max(0, Number(document.getElementById("percentualMeta").value) || 0));

    let totalReferencia = 0;
    let concluidasReferencia = 0;
    let descricaoMeta = "";

    if (tipoVisualizacao === "ano") {
        totalReferencia = dataFiltrada.length;
        concluidasReferencia = concluidasFiltradas;
        descricaoMeta = "Meta anual";
    }

    if (tipoVisualizacao === "mes") {
        totalReferencia = dataFiltrada.length;
        concluidasReferencia = concluidasFiltradas;
        descricaoMeta = "Meta mensal";
    }

    if (tipoVisualizacao === "dia") {
        let dadosMes = obterDadosDoMesSelecionado();

        totalReferencia = dadosMes.length;
        concluidasReferencia = dadosMes.filter(row => estaConcluida(row)).length;
        descricaoMeta = "Meta mensal";
    }

    let metaReferencia = Math.round((totalReferencia * percentual) / 100);

    let progresso = 0;

    if (metaReferencia > 0) {
        progresso = Math.round((concluidasReferencia / metaReferencia) * 100);
    } else {
        progresso = totalReferencia === 0 ? 0 : 100;
    }

    if (progresso > 100) progresso = 100;

    let falta = Math.max(0, 100 - progresso);

    document.getElementById("metaTexto").innerText =
        `${descricaoMeta}: ${metaReferencia} | Concluídas: ${concluidasReferencia}`;

    document.getElementById("metaProgressBar").style.width = `${progresso}%`;

    if (falta === 0) {
        document.getElementById("metaProgressText").innerText = `Meta atingida`;
        document.getElementById("metaProgressText").style.color = "#28a745";
    } else {
        document.getElementById("metaProgressText").innerText = `Faltam ${falta}% para atingir a meta`;
        document.getElementById("metaProgressText").style.color = "#dc3545";
    }
}
/* PROCURAR */
function procurarManutencao() {
    termoProcuraAtivo = document.getElementById("campoProcura").value;
    atualizarListaManutencao();
}
/* VERIFICAR PROCURA DA LISTA */
function passaFiltrosLista(row) {
    let procura = normalizarTexto(termoProcuraAtivo);

    let valores = {
        data: formatarData(getData(row)),
        inicio: formatarDataHora(getHoraInicio(row)),
        fim: formatarDataHora(getHoraConclusao(row)),
        equipamento: getEquipamento(row) || "",
        setor: getSetor(row) || "",
        tipo: getTipo(row) || "",
        falha: getFalha(row) || "",
        estado: getEstado(row) || "",
        conclusao: formatarData(getDataConclusao(row))
    };

    if (procura) {
        let textoLinha = normalizarTexto(Object.values(valores).join(" "));
        if (!textoLinha.includes(procura)) return false;
    }

    return true;
}
/* ORDENAÇÃO DA LISTA */
function ordenarLista(coluna) {
    if (ordenacaoAtual.coluna === coluna) {
        ordenacaoAtual.direcao = ordenacaoAtual.direcao === "asc" ? "desc" : "asc";
    } else {
        ordenacaoAtual.coluna = coluna;
        ordenacaoAtual.direcao = "asc";
    }

    atualizarListaManutencao();
    atualizarHistoricoManutencao();
}

function valorOrdenacao(row, coluna) {
    
    if (coluna === "insercao") return excelDateToJSDate(getHoraInicio(row))?.getTime() || 0;
    if (coluna === "conclusao") return excelDateToJSDate(getDataConclusao(row))?.getTime() || 0;
    if (coluna === "equipamento") return normalizarTexto(getEquipamento(row));
    if (coluna === "setor") return normalizarTexto(getSetor(row));
    if (coluna === "tipo") return normalizarTexto(getTipo(row));
    if (coluna === "falha") return normalizarTexto(getFalha(row));
    if (coluna === "estado") return normalizarTexto(getEstado(row));

    return "";
}

function ordenarItens(lista) {
    if (!ordenacaoAtual.coluna) return lista;

    return lista.sort((a, b) => {
        let va = valorOrdenacao(a.row, ordenacaoAtual.coluna);
        let vb = valorOrdenacao(b.row, ordenacaoAtual.coluna);

        if (typeof va === "number" && typeof vb === "number") {
            return ordenacaoAtual.direcao === "asc" ? va - vb : vb - va;
        }

        return ordenacaoAtual.direcao === "asc"
            ? String(va).localeCompare(String(vb), "pt")
            : String(vb).localeCompare(String(va), "pt");
    });
}

function aplicarCorAtraso(tr, row) {
    if (estaConcluida(row)) return;

    let dias = diferencaDias(getData(row), new Date());

    if (dias > 15) {
        tr.classList.add("linha-vermelha");
        return;
    }

    if (dias > 7) {
        tr.classList.add("linha-laranja");
        return;
    }

    if (dias >= 2) {
        tr.classList.add("linha-amarela");
    }
}

function adicionarCelulaTexto(tr, valor) {
    const td = document.createElement("td");
    td.textContent = valor === null || valor === undefined ? "" : String(valor);
    tr.appendChild(td);
    return td;
}

function criarLinhaTabela(row, index) {
    const tr = document.createElement("tr");
    tr.classList.add("row-click");
    tr.onclick = () => {
        selecionado = row;
        document.querySelectorAll("#tabelaManutencao tbody tr").forEach(r => r.classList.remove("linha-selecionada"));
        tr.classList.add("linha-selecionada");
        document.getElementById("btnEditar").style.display = "block";
        document.getElementById("btnDeletar").style.display = "block";
    };
    aplicarCorAtraso(tr, row);
    [formatarDataHora(getHoraInicio(row) || getData(row)), formatarDataHora(getHoraConclusao(row) || getDataConclusao(row)), getManutencao(row), getEquipamento(row), getSetor(row), getTipo(row), getFalha(row), getEstado(row)].forEach(v => adicionarCelulaTexto(tr, v));
    return tr;
}
/* LISTA MANUTENÇÕES PENDENTES */
function atualizarListaManutencao() {
    let tbody = document.querySelector("#tabelaManutencao tbody");
    tbody.innerHTML = "";

    let lista = dadosOriginais
        .map((row, index) => ({ row, index }))
        .filter(item => passaFiltrosLista(item.row))
        .filter(item => !estaConcluida(item.row));

    lista = ordenarItens(lista);

    lista.forEach(item => {
        tbody.appendChild(criarLinhaTabela(item.row, item.index));
    });
}
/* HISTÓRICO DE MANUTENÇÕES CONCLUÍDAS */
function atualizarHistoricoManutencao() {

    let tbody = document.querySelector("#tabelaHistorico tbody");
    tbody.innerHTML = "";

    let lista = dadosOriginais
        .map((row, index) => ({ row, index }))

        // ✅ filtrar concluídos
        .filter(item => (getEstado(item.row) || "").toLowerCase().includes("conclu"));

    // ✅ AQUI ESTÁ O QUE FALTAVA
    lista = ordenarItens(lista);

    lista.forEach(item => {

        let row = item.row;

        let tr = document.createElement("tr");

        tr.onclick = () => {

            selecionado = row;

            document.querySelectorAll("#tabelaHistorico tbody tr")
                .forEach(r => r.classList.remove("linha-selecionada"));

            tr.classList.add("linha-selecionada");

            document.getElementById("btnEditarHistorico").style.display = "block";
            document.getElementById("btnDeletarHistorico").style.display = "block";

            console.log("Selecionado histórico:", selecionado);
        };

        [formatarDataHora(getHoraInicio(row) || getData(row)), formatarDataHora(getHoraConclusao(row) || getDataConclusao(row)), getManutencao(row), getEquipamento(row), getSetor(row), getTipo(row), getFalha(row), getEstado(row)].forEach(v => adicionarCelulaTexto(tr, v));

        tbody.appendChild(tr);
    });
}

function preencherSelect(id, lista, textoInicial) {
    let select = document.getElementById(id);
    select.innerHTML = "";

    let opt = document.createElement("option");
    opt.value = "";
    opt.textContent = textoInicial;
    select.appendChild(opt);

    lista.sort().forEach(item => {
        let option = document.createElement("option");
        option.value = item;
        option.textContent = item;
        select.appendChild(option);
    });
}
/* NOVA MANUTENÇÃO */
function abrirNovaManutencao() {
    if (paginaAtual !== "manutencao" && paginaAtual !== "historico") return;
    document.getElementById("painelManutencao").style.display = "block";
    atualizarDropdownEquipamentos();
    atualizarDropdownSetores();
}
function guardarNovaManutencao() {
    try {
        const manutencao = document.getElementById("novoManutencao").value.trim();
        const equipamentoNome = document.getElementById("novoEquipamento").value;
        const setor = document.getElementById("novoSetor").value;
        const tipo = document.getElementById("novoTipo").value;
        const falha = document.getElementById("novoFalha")?.value || "";
        const frequencia = document.getElementById("novoFrequencia").value;
        const estado = document.getElementById("novoEstado").value || "Pendente";
        const descricao = document.getElementById("novoDescricao").value;
        const inicioTexto = document.getElementById("novoHoraInicio").value;
        const conclusaoTexto = document.getElementById("novoHoraConclusao").value;
        const inicio = inicioTexto ? excelDateToJSDate(inicioTexto) : new Date();
        const conclusao = conclusaoTexto ? excelDateToJSDate(conclusaoTexto) : "";

        if (!equipamentoNome) return mostrarMensagem("Selecione um equipamento");
        if (!inicio) return mostrarMensagem("A data de início é inválida");
        if (conclusaoTexto && !conclusao) return mostrarMensagem("A data de conclusão é inválida");
        if (conclusao && conclusao < inicio) return mostrarMensagem("A conclusão não pode ser anterior ao início");

        const campos = {
            "Hora de Início": inicio,
            "Hora de Conclusão": conclusao,
            "Manutenção": manutencao,
            "Equipamento": equipamentoNome,
            "Setor": setor,
            "Tipo de Manutenção": tipo,
            "Frequencia": frequencia,
            "Tipo de Falha": falha,
            "Descricao": descricao,
            "Estado": estado
        };
        if (selecionado && dadosOriginais.includes(selecionado)) {
            // Preserva colunas adicionais que existam no Excel.
            Object.assign(selecionado, campos);
            mostrarMensagem("Manutenção editada");
        } else {
            dadosOriginais.push(campos);
            selecionado = null;
            mostrarMensagem("Manutenção criada");
        }
        if (!guardarDadosLocalmente()) return;
        guardarEstoque();
        guardarEquipamentos();
        preencherFiltroAno(dadosOriginais);
        atualizarFiltrosDashboard();
        atualizarListaManutencao();
        atualizarHistoricoManutencao();
        fecharPainelManutencao();
        limparFormularioManutencao();
        ["btnEditar", "btnDeletar", "btnEditarHistorico", "btnDeletarHistorico"].forEach(id => { const b=document.getElementById(id); if (b) b.style.display="none"; });
    } catch (erro) {
        console.error("Erro ao guardar manutenção:", erro);
        mostrarMensagem("Erro ao guardar manutenção");
    }
}
function concluirManutencao() {
    const row = selecionado;
    if (!row || !dadosOriginais.includes(row)) return mostrarMensagem("Selecione uma manutenção válida");
    const inicio = excelDateToJSDate(getHoraInicio(row) || getData(row));
    if (!inicio) return mostrarMensagem("Defina uma data de início válida antes de concluir");
    const agora = new Date();
    definirCampo(row, ["estado"], "Estado", "Concluído");
    definirCampo(row, ["data de conclusao", "data de conclusão", "conclusao", "conclusão"], "Data de Conclusão", agora);
    definirCampo(row, ["hora de conclusao", "hora de conclusão", "hora conclusao", "hora conclusão", "fim", "fim defeito", "hora final"], "Hora de Conclusão", agora);
    if (normalizarTexto(getTipo(row)).includes("preventiva")) criarProximaPreventivaRecorrente(row);
    guardarDadosLocalmente();
    atualizarFiltrosDashboard();
    atualizarListaManutencao();
    atualizarHistoricoManutencao();
    mostrarMensagem("Manutenção concluída. Exporte o Excel para criar uma cópia externa.");
}
function deletarSelecionado() {

    if (!selecionado) {
        return mostrarMensagem("Selecione uma manutenção");
    }

    if (!confirm("Deseja deletar esta manutenção?")) return;

    let index = dadosOriginais.indexOf(selecionado);

    if (index !== -1) {
        dadosOriginais.splice(index, 1);
    }

    selecionado = null;

    guardarDadosLocalmente();
    atualizarListaManutencao();
    atualizarHistoricoManutencao();

    // esconder botões histórico
    let btnE = document.getElementById("btnEditarHistorico");
    let btnD = document.getElementById("btnDeletarHistorico");

    if (btnE) btnE.style.display = "none";
    if (btnD) btnD.style.display = "none";

    mostrarMensagem("🗑️ Removido com sucesso");
}
/* EDITAR */
function editarManutencao() {
    const row = selecionado;
    if (!row || !dadosOriginais.includes(row)) return mostrarMensagem("Selecione uma manutenção");
    mostrarPagina("manutencao");
    abrirNovaManutencao();
    atualizarDropdownEquipamentos();
    atualizarDropdownSetores();
    document.getElementById("novoManutencao").value = getManutencao(row) || "";
    document.getElementById("novoEquipamento").value = getEquipamento(row) || "";
    document.getElementById("novoSetor").value = getSetor(row) || "";
    document.getElementById("novoTipo").value = getTipo(row) || "";
    document.getElementById("novoFalha").value = getFalha(row) || "";
    document.getElementById("novoFrequencia").value = getFrequencia(row) || "";
    document.getElementById("novoEstado").value = getEstado(row) || "Pendente";
    document.getElementById("novoDescricao").value = getDescricao(row) || "";
    document.getElementById("novoHoraInicio").value = formatarDataHoraInput(getHoraInicio(row) || getData(row));
    document.getElementById("novoHoraConclusao").value = formatarDataHoraInput(getHoraConclusao(row) || getDataConclusao(row));
}
/* LIMPAR */
function limparFormularioManutencao() {

    let ids = [
        "novoManutencao",
        "novoEquipamento",
        "novoSetor",
        "novoFalha",
        "novoTipo",
        "novoFrequencia",
        "novoEstado",
        "novoDescricao",
        "novoHoraInicio",
        "novoHoraConclusao"
    ];

    ids.forEach(id => {
        let el = document.getElementById(id);
        if (el) el.value = "";
    });
}


function atualizarTabelaEstoque() {
    const tbody = document.querySelector("#tabelaEstoque tbody");
    if (!tbody) return;
    tbody.replaceChildren();
    estoque.forEach((item, index) => {
        const tr = document.createElement("tr");
        tr.onclick = () => {
            itemSelecionado = index;
            tbody.querySelectorAll("tr").forEach(r => r.classList.remove("linha-selecionada"));
            tr.classList.add("linha-selecionada");
        };
        [item.nome, item.serie, item.equipamento, item.local, item.quantidade].forEach(v => adicionarCelulaTexto(tr, v));
        tbody.appendChild(tr);
    });
}
function abrirNovoItem() {
    itemSelecionado = null;
    document.getElementById("itemNome").value = "";
    document.getElementById("itemSerie").value = "";
    document.getElementById("itemEquip").value = "";
    document.getElementById("itemLocal").value = "";
    document.getElementById("itemQtd").value = "";
    mostrarPagina("novoItem");
}

function guardarItem(e) {
    if (e) e.preventDefault();
    const nome = document.getElementById("itemNome").value.trim();
    const quantidadeTexto = document.getElementById("itemQtd").value;
    const quantidade = Number(quantidadeTexto);
    if (!nome) return mostrarMensagem("Indique o nome do item");
    if (quantidadeTexto === "" || !Number.isFinite(quantidade) || quantidade < 0) return mostrarMensagem("Indique uma quantidade válida (zero ou superior)");
    const novo = {
        nome,
        serie: document.getElementById("itemSerie").value.trim(),
        equipamento: document.getElementById("itemEquip").value.trim(),
        local: document.getElementById("itemLocal").value.trim(),
        quantidade
    };
    if (itemSelecionado !== null && estoque[itemSelecionado]) {
        estoque[itemSelecionado] = novo;
        mostrarMensagem("Item editado");
    } else {
        estoque.push(novo);
        mostrarMensagem("Item adicionado");
    }
    itemSelecionado = null;
    guardarEstoque();
    atualizarTabelaEstoque();
    mostrarPagina("estoque");
}

function editarItem() {
    if (itemSelecionado === null || !estoque[itemSelecionado]) return mostrarMensagem("Selecione um item");
    const item = estoque[itemSelecionado];
    document.getElementById("itemNome").value = item.nome || "";
    document.getElementById("itemSerie").value = item.serie || "";
    document.getElementById("itemEquip").value = item.equipamento || "";
    document.getElementById("itemLocal").value = item.local || "";
    document.getElementById("itemQtd").value = item.quantidade ?? 0;
    mostrarPagina("novoItem");
}
function deletarItem() {

    if (itemSelecionado === null)
        return mostrarMensagem("Selecione um item");

    if (!confirm("Deseja deletar este item?")) return;

    estoque.splice(itemSelecionado, 1);

    itemSelecionado = null;

    guardarEstoque();
    atualizarTabelaEstoque();
    mostrarMensagem("Item removido");
}
// =============================
// ✅ EXPORTAÇÃO AUTOMÁTICA EXCEL
// =============================
function exportarExcelAutomatico() {

    let wb = XLSX.utils.book_new();

    // ✅ MANUTENÇÃO
    let ws1 = XLSX.utils.json_to_sheet(dadosOriginais);
    XLSX.utils.book_append_sheet(wb, ws1, "Manutenção");

    // ✅ ESTOQUE
    let dadosEstoque = estoque.map(item => ({
    "Item": item.nome,
    "Número Série": item.serie,
    "Equipamento": item.equipamento,
    "Local": item.local,
    "Quantidade": item.quantidade
    }));

    let ws2 = XLSX.utils.json_to_sheet(dadosEstoque);
    XLSX.utils.book_append_sheet(wb, ws2, "Estoque");

    // ✅ EQUIPAMENTOS (NOVO)
    let ws3 = XLSX.utils.json_to_sheet(equipamentos);
    XLSX.utils.book_append_sheet(wb, ws3, "Equipamentos");

    XLSX.writeFile(wb, "dados_manutencao.xlsx");

    console.log("✅ Excel exportado com equipamentos");
}
// =============================
// ✅ IMPORTAR EXCEL
// =============================
function importarExcel(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onerror = () => mostrarMensagem("Não foi possível ler o ficheiro Excel");
    reader.onload = async function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const nomeAbaManutencao = workbook.SheetNames.find(n => normalizarTexto(n).includes("manutencao")) || workbook.SheetNames[0];
            const sheet = workbook.Sheets[nomeAbaManutencao];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            if (!json.length) return mostrarMensagem("Excel vazio ou sem registos de manutenção");

            const novosDados = json.map(original => {
                const row = { ...original };
                const inicio = excelDateToJSDate(getData(row));
                const conclusao = excelDateToJSDate(getDataConclusao(row));
                const estado = String(getEstado(row) || "Pendente").trim();
                // Preserva todas as colunas originais e uniformiza os campos usados pela aplicação.
                return {
                    ...row,
                    "Hora de Início": excelDateToJSDate(getHoraInicio(row)) || inicio || null,
                    "Hora de Conclusão": excelDateToJSDate(getHoraConclusao(row)) || conclusao || "",
                    "Manutenção": String(getManutencao(row) || ""),
                    "Equipamento": String(getEquipamento(row) || ""),
                    "Setor": String(getSetor(row) || "").replace(/,+$/, "").trim(),
                    "Tipo de Manutenção": String(getTipo(row) || ""),
                    "Tipo de Falha": String(getFalha(row) || ""),
                    "Estado": estado
                };
            });

            let novosEquipamentos = equipamentos;
            const sheetEquip = workbook.SheetNames.find(n => normalizarTexto(n).includes("equip"));
            if (sheetEquip) {
                novosEquipamentos = XLSX.utils.sheet_to_json(workbook.Sheets[sheetEquip], { defval: "" })
                    .map(eq => ({ nome: String(eq.nome || eq.Equipamento || "").trim(), setor: String(eq.setor || eq.Setor || "").trim() }))
                    .filter(eq => eq.nome);
            }
            let novoEstoque = estoque;
            const sheetEstoque = workbook.SheetNames.find(n => normalizarTexto(n).includes("estoque"));
            if (sheetEstoque) {
                novoEstoque = XLSX.utils.sheet_to_json(workbook.Sheets[sheetEstoque], { defval: "" }).map(item => {
                    const raw = item["Quantidade"];
                    const qtd = raw === "" || raw === null ? 0 : Number(raw);
                    return { nome: String(item["Item"] || ""), serie: String(item["Número Série"] || item["Numero Série"] || ""), equipamento: String(item["Equipamento"] || ""), local: String(item["Local"] || ""), quantidade: Number.isFinite(qtd) && qtd >= 0 ? qtd : 0 };
                });
            }

            // Só atualizar o estado depois de analisar o ficheiro inteiro.
            dadosOriginais = novosDados;
            equipamentos = novosEquipamentos;
            estoque = novoEstoque;
            normalizarHorariosExistentes();
            guardarDadosLocalmente();
            guardarEquipamentos();
            guardarEstoque();
            preencherFiltroAno(dadosOriginais);
            atualizarFiltrosDashboard();
            atualizarListaManutencao();
            atualizarHistoricoManutencao();
            atualizarTabelaEquipamentos();
            atualizarTabelaEstoque();
            atualizarDropdownEquipamentos();
            atualizarDropdownSetores();
            mostrarPagina("dashboard");

            // A cópia do Excel é opcional: não deixar uma falha de quota desfazer a importação.
            try {
                let binary = "";
                const chunkSize = 0x8000;
                for (let i = 0; i < data.length; i += chunkSize) binary += String.fromCharCode(...data.subarray(i, i + chunkSize));
                localStorage.setItem(CHAVE_ULTIMO_EXCEL, btoa(binary));
            } catch (storageError) {
                console.warn("Não foi possível guardar a cópia do Excel:", storageError);
            }
            mostrarMensagem(`Importação concluída: ${dadosOriginais.length} registos de manutenção`);
        } catch (erro) {
            console.error("Erro ao importar Excel:", erro);
            mostrarMensagem("Falha na importação. Os dados anteriores foram mantidos quando possível.");
        }
    };
    reader.readAsArrayBuffer(file);
}

async function carregarUltimoExcel() {
    const base64 = localStorage.getItem(CHAVE_ULTIMO_EXCEL);
    if (!base64) return;
    try {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const file = new File([bytes], "ultimo_excel.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        importarExcel(file);
    } catch (erro) {
        console.error("Erro ao recuperar último Excel:", erro);
        mostrarMensagem("Não foi possível recuperar a última cópia do Excel");
    }
}

async function carregarExcelOnline() {
    const url = "https://atrya365-my.sharepoint.com/:x:/g/personal/jelias_acwin_pt/IQDYMVKMgblkQ5Z__OVS9butAbOwoQihOzzphANEuQ3QMys?e=Zgjerw";
    try {
        const resposta = await fetch(url);
        if (!resposta.ok) throw new Error(`HTTP ${resposta.status}`);
        const data = await resposta.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const nomeAba = workbook.SheetNames.find(n => normalizarTexto(n).includes("manutencao")) || workbook.SheetNames[0];
        const sheet = workbook.Sheets[nomeAba];
        if (!sheet) throw new Error("Aba de manutenção não encontrada");
        // Reutilizar a mesma rotina de importação para manter o mapeamento consistente.
        importarExcel(new File([data], "manutencao_online.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    } catch (erro) {
        console.error("Erro ao carregar Excel online:", erro);
        mostrarMensagem("Não foi possível carregar o Excel online. Importe um ficheiro manualmente.");
    }
}

function mostrarMensagem(texto) {

    let box = document.getElementById("mensagemSistema");

    box.innerText = texto;
    box.style.display = "block";

    // esconder depois de 3 segundos
    setTimeout(() => {
        box.style.display = "none";
    }, 3000);
}

function atualizarEstadoManual() {

    let estado = document.getElementById("novoEstado").value;
    let campoConclusao = document.getElementById("novoHoraConclusao");

    // se marcar como concluído e não tiver hora → preenche
    if (estado === "Concluído" && !campoConclusao.value) {
        let agora = new Date();
        campoConclusao.value = agora.toISOString().slice(0,16);
    }
}


function atualizarPreviewEstado() {
    // não faz nada crítico → só existe para não quebrar
    return;
}

function guardarEstoque() {
    try { localStorage.setItem("estoque_dados", JSON.stringify(estoque)); return true; }
    catch (erro) { console.error("Erro ao guardar stock:", erro); mostrarMensagem("Não foi possível guardar o stock localmente"); return false; }
}

function carregarEstoque() {
    const dados = localStorage.getItem("estoque_dados");
    if (!dados) return;
    try { const parsed = JSON.parse(dados); if (Array.isArray(parsed)) estoque = parsed; }
    catch (erro) { console.error("Erro ao carregar stock:", erro); }
}

function guardarEquipamentos() {
    try { localStorage.setItem("equipamentos", JSON.stringify(equipamentos)); return true; }
    catch (erro) { console.error("Erro ao guardar equipamentos:", erro); mostrarMensagem("Não foi possível guardar os equipamentos localmente"); return false; }
}

function carregarEquipamentos() {
    const dados = localStorage.getItem("equipamentos");
    if (!dados) return;
    try { const parsed = JSON.parse(dados); if (Array.isArray(parsed)) equipamentos = parsed; }
    catch (erro) { console.error("Erro ao carregar equipamentos:", erro); }
}

function atualizarTabelaEquipamentos() {
    const tbody = document.querySelector("#tabelaEquipamentos tbody");
    if (!tbody) return;
    tbody.replaceChildren();
    equipamentos.forEach((item, index) => {
        const tr = document.createElement("tr");
        tr.onclick = () => {
            equipamentoSelecionado = index;
            tbody.querySelectorAll("tr").forEach(r => r.classList.remove("linha-selecionada"));
            tr.classList.add("linha-selecionada");
        };
        [item.nome, item.setor].forEach(v => adicionarCelulaTexto(tr, v));
        tbody.appendChild(tr);
    });
}
function abrirModalEquipamento() {

    modoEdicaoEquip = false;

    document.getElementById("equipNome").value = "";
    document.getElementById("equipSetor").value = "";

    document.getElementById("modalEquipamento").style.display = "block";
}
function editarEquipamento() {

    if (equipamentoSelecionado === null)
        return mostrarMensagem("Selecione um equipamento");

    let item = equipamentos[equipamentoSelecionado];

    modoEdicaoEquip = true;

    document.getElementById("equipNome").value = item.nome;
    document.getElementById("equipSetor").value = item.setor;

    document.getElementById("modalEquipamento").style.display = "block";
    
    atualizarDropdownEquipamentos();

}
function guardarEquipamento() {

    let nome = document.getElementById("equipNome").value.trim();
    let setor = document.getElementById("equipSetor").value.trim();

    if (!nome || !setor)
        return mostrarMensagem("Preencha todos os campos");

    if (modoEdicaoEquip) {

        equipamentos[equipamentoSelecionado] = { nome, setor };

        mostrarMensagem("✅ Equipamento editado");

    } else {

        equipamentos.push({ nome, setor });

        mostrarMensagem("✅ Equipamento adicionado");
    }

    guardarEquipamentos();
    atualizarTabelaEquipamentos();
    atualizarDropdownEquipamentos();

    fecharModalEquipamento();

    equipamentoSelecionado = null;
}
function deletarEquipamento() {

    if (equipamentoSelecionado === null)
        return mostrarMensagem("Selecione um equipamento");

    if (!confirm("Deseja apagar este equipamento?")) return;

    equipamentos.splice(equipamentoSelecionado, 1);

    guardarEquipamentos();
    atualizarTabelaEquipamentos();
    atualizarDropdownEquipamentos();

    equipamentoSelecionado = null;
    
    mostrarMensagem("🗑️ Equipamento removido");
}

function fecharModalEquipamento() {
    document.getElementById("modalEquipamento").style.display = "none";
}

function atualizarDropdownSetores() {

    let selects = [
        document.getElementById("novoSetor"),
        document.getElementById("editSetor")
    ];

    selects.forEach(select => {

        if (!select) return;

        let valorAtual = select.value;

        select.innerHTML = '<option value="">-- Selecionar --</option>';

        setores.forEach(setor => {
            let opt = document.createElement("option");
            opt.value = setor;
            opt.textContent = setor;
            select.appendChild(opt);
        });

        select.value = valorAtual;
    });
}

function atualizarDropdownEquipamentos() {

    let selects = [
        document.getElementById("novoEquipamento"),
        document.getElementById("editEquipamento")
    ];

    selects.forEach(select => {

        if (!select) return;

        let valorAtual = select.value; // ✅ mantém seleção (se existir)

        select.innerHTML = '<option value="">-- Selecionar --</option>';

        equipamentos.forEach(eq => {

            let opt = document.createElement("option");

            opt.value = eq.nome;
            opt.textContent = `${eq.nome} (${eq.setor})`;

            select.appendChild(opt);
        });

        // ✅ tenta restaurar valor (edição)
        select.value = valorAtual;
    });
}

function fecharPainelManutencao() {
    document.getElementById("painelManutencao").style.display = "none";
}
