// =====================================================
// CONCILIAÇÃO BANCÁRIA - app.js (versão completa)
// =====================================================

const estado = {
  extrato: [],
  fornecedores: [],
  conciliados: [],
  pendentes: [],
  lancamentos: [],
  historico: JSON.parse(localStorage.getItem('conciliacao_historico') || '[]'),
  itemEditando: null
};

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const telaUpload     = $('#tela-upload');
const telaResultado  = $('#tela-resultado');
const telaHistorico  = $('#tela-historico');
const fileExtrato    = $('#file-extrato');
const fileFornecedores = $('#file-fornecedores');
const btnConciliar   = $('#btn-conciliar');
const infoExtrato    = $('#info-extrato');
const infoFornecedores = $('#info-fornecedores');

// -------------------- NAVEGAÇÃO --------------------
function mostrarTela(tela) {
  $$('.tela').forEach(t => t.classList.remove('ativa'));
  tela.classList.add('ativa');
}

$('#btn-nova').addEventListener('click', () => {
  resetarUpload();
  mostrarTela(telaUpload);
});

$('#btn-historico').addEventListener('click', () => {
  renderizarHistorico();
  mostrarTela(telaHistorico);
});

$('#btn-voltar').addEventListener('click', () => {
  mostrarTela(telaUpload);
});

// -------------------- UPLOAD --------------------
fileExtrato.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  lerArquivo(file, (dados) => {
    try {
      estado.extrato = normalizarExtrato(dados);
      console.log('Extrato carregado:', estado.extrato.length, 'linhas');
      console.table(estado.extrato.slice(0, 5));
      mostrarInfoArquivo(infoExtrato, file.name);
      $('#card-extrato').classList.add('has-file');
      verificarPodeConciliar();
    } catch (err) {
      console.error('Erro ao processar extrato:', err);
      alert('Erro ao processar o extrato.');
    }
  });
});

fileFornecedores.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  lerArquivo(file, (dados) => {
    try {
      estado.fornecedores = normalizarFornecedores(dados);
      console.log('Fornecedores carregados:', estado.fornecedores.length, 'linhas');
      console.table(estado.fornecedores.slice(0, 5));
      mostrarInfoArquivo(infoFornecedores, file.name);
      $('#card-fornecedores').classList.add('has-file');
      verificarPodeConciliar();
    } catch (err) {
      console.error('Erro ao processar fornecedores:', err);
      alert('Erro ao processar o relatório de fornecedores.');
    }
  });
});

infoExtrato.querySelector('.btn-remove').addEventListener('click', () => {
  estado.extrato = [];
  fileExtrato.value = '';
  infoExtrato.classList.add('hidden');
  $('#card-extrato').classList.remove('has-file');
  verificarPodeConciliar();
});

infoFornecedores.querySelector('.btn-remove').addEventListener('click', () => {
  estado.fornecedores = [];
  fileFornecedores.value = '';
  infoFornecedores.classList.add('hidden');
  $('#card-fornecedores').classList.remove('has-file');
  verificarPodeConciliar();
});

function mostrarInfoArquivo(el, nome) {
  el.querySelector('.file-name').textContent = nome;
  el.classList.remove('hidden');
}

function verificarPodeConciliar() {
  btnConciliar.disabled = !(estado.extrato.length && estado.fornecedores.length);
}

function resetarUpload() {
  estado.extrato = [];
  estado.fornecedores = [];
  estado.conciliados = [];
  estado.pendentes = [];
  estado.lancamentos = [];
  fileExtrato.value = '';
  fileFornecedores.value = '';
  infoExtrato.classList.add('hidden');
  infoFornecedores.classList.add('hidden');
  $('#card-extrato').classList.remove('has-file');
  $('#card-fornecedores').classList.remove('has-file');
  btnConciliar.disabled = true;
}

// -------------------- LEITURA DE ARQUIVO (CSV + Excel) --------------------
function lerArquivo(file, callback) {
  const nome = file.name.toLowerCase();
  const isExcel = nome.endsWith('.xlsx') || nome.endsWith('.xls');

  if (isExcel) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const primeiraAba = workbook.SheetNames[0];
        const planilha = workbook.Sheets[primeiraAba];
        const json = XLSX.utils.sheet_to_json(planilha, { defval: '' });
        console.log('Excel lido com sucesso. Linhas:', json.length);
        callback(json);
      } catch (err) {
        console.error('Erro ao ler Excel:', err);
        alert('Erro ao ler o arquivo Excel.');
      }
    };
    reader.onerror = () => alert('Erro ao ler o arquivo.');
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        let texto = e.target.result;
        if (texto.charCodeAt(0) === 0xFEFF) texto = texto.slice(1);
        const linhas = parseCSV(texto);
        callback(linhas);
      } catch (err) {
        console.error('Erro ao ler CSV:', err);
        alert('Não foi possível ler o arquivo CSV.');
      }
    };
    reader.onerror = () => alert('Erro ao ler o arquivo.');
    reader.readAsText(file, 'UTF-8');
  }
}

function parseCSV(texto) {
  const linhas = texto.trim().split(/\r?\n/);
  if (linhas.length < 2) return [];

  const primeira = linhas[0];
  const separador = (primeira.match(/;/g) || []).length >= (primeira.match(/,/g) || []).length ? ';' : ',';

  const headers = primeira.split(separador).map(h => h.trim().toLowerCase().replace(/['"]/g, ''));
  const dados = [];

  for (let i = 1; i < linhas.length; i++) {
    const linha = linhas[i].trim();
    if (!linha) continue;
    const valores = linha.split(separador).map(v => v.trim().replace(/['"]/g, ''));
    const obj = {};
    headers.forEach((h, idx) => obj[h] = valores[idx] || '');
    dados.push(obj);
  }
  return dados;
}

// -------------------- NORMALIZAÇÃO (versão agressiva) --------------------
function normalizarExtrato(dados) {
  console.log('=== DEBUG EXTRATO ===');
  console.log('Linhas recebidas:', dados.length);
  if (dados.length > 0) {
    console.log('Colunas do Extrato:', Object.keys(dados[0]));
  }

  const resultado = dados.map((linha, idx) => {
    const data = encontrarValor(linha, [
      'data', 'date', 'dt', 'data lançamento', 'data lancamento', 'data movimentação'
    ]);
    
    const documento = encontrarValor(linha, [
      'documento', 'doc', 'nº documento', 'numero', 'nro', 'id', 'n doc', 'núm. documento'
    ]);
    
    const descricao = encontrarValor(linha, [
      'descricao', 'descrição', 'historico', 'histórico', 'memo', 'nome', 
      'lançamento', 'lancamento', 'histórico completo', 'complemento'
    ]);
    
    const { valor, tipo } = detectarValorTipo(linha);

    return {
      id: `ext-${idx}`,
      origem: 'extrato',
      data: formatarData(data),
      documento: String(documento || '').replace(/\D/g, ''),
      descricao: descricao || '',
      valor: valor,
      tipo: tipo, // 'debito' | 'credito' | null
      usado: false,
      raw: linha
    };
  });

  const comValor = resultado.filter(i => i.valor > 0);
  console.log('Extrato com valor > 0:', comValor.length);

  // Se nenhum tiver valor, usa todos mesmo assim (para não travar)
  if (comValor.length === 0 && resultado.length > 0) {
    console.warn('Nenhum valor encontrado no extrato. Usando todas as linhas.');
    return resultado;
  }

  return comValor;
}

function normalizarFornecedores(dados) {
  console.log('=== DEBUG FORNECEDORES ===');
  console.log('Linhas recebidas:', dados.length);
  if (dados.length > 0) {
    console.log('Colunas do Fornecedores:', Object.keys(dados[0]));
  }

  const resultado = dados.map((linha, idx) => {
    const data = encontrarValor(linha, [
      'data', 'date', 'dt', 'data pagamento', 'data pagto', 'data emissão'
    ]);
    
    const documento = encontrarValor(linha, [
      'documento', 'doc', 'nº documento', 'numero', 'nro', 'nf', 'nota', 'n doc', 'número nf'
    ]);
    
    const fornecedor = encontrarValor(linha, [
      'fornecedor', 'nome', 'razao social', 'razão social', 'participante', 
      'cliente', 'favorecido', 'beneficiário', 'nome fantasia'
    ]);
    
    const valor = parseValor(encontrarValor(linha, [
      'valor', 'value', 'amount', 'vlr', 'valor pago', 'valor (r$)', 
      'valor total', 'total', 'líquido'
    ]));

    return {
      id: `forn-${idx}`,
      origem: 'fornecedor',
      data: formatarData(data),
      documento: String(documento || '').replace(/\D/g, ''),
      fornecedor: fornecedor || '',
      valor: valor,
      usado: false,
      raw: linha
    };
  });

  const comValor = resultado.filter(i => i.valor > 0);
  console.log('Fornecedores com valor > 0:', comValor.length);

  if (comValor.length === 0 && resultado.length > 0) {
    console.warn('Nenhum valor encontrado nos fornecedores. Usando todas as linhas.');
    return resultado;
  }

  return comValor;
}

function encontrarValor(obj, chaves) {
  const keys = Object.keys(obj);
  
  for (const chave of chaves) {
    const encontrada = keys.find(k => {
      const kLimpo = k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      const cLimpo = chave.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return kLimpo.includes(cLimpo);
    });
    if (encontrada !== undefined) return obj[encontrada];
  }
  return '';
}

function parseValorComSinal(str) {
  if (str === null || str === undefined || str === '') return 0;
  let limpo = String(str)
    .replace(/R\$\s?/gi, '')
    .replace(/\s/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const num = parseFloat(limpo);
  return isNaN(num) ? 0 : num;
}

function parseValor(str) {
  return Math.abs(parseValorComSinal(str));
}

// Detecta se a linha do extrato é um débito (saída) ou crédito (entrada),
// para não conciliar pagamentos a fornecedores com depósitos/entradas.
function detectarValorTipo(linha) {
  const keys = Object.keys(linha);
  const norm = k => k.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  const chaveDebito = keys.find(k => /d[ée]bito|sa[íi]da/.test(norm(k)));
  const chaveCredito = keys.find(k => /cr[ée]dito|entrada/.test(norm(k)));

  // Colunas separadas de débito/crédito: usamos elas, é a informação mais confiável
  if (chaveDebito !== undefined || chaveCredito !== undefined) {
    const vDebito = chaveDebito !== undefined ? Math.abs(parseValorComSinal(linha[chaveDebito])) : 0;
    const vCredito = chaveCredito !== undefined ? Math.abs(parseValorComSinal(linha[chaveCredito])) : 0;
    if (vDebito > 0) return { valor: vDebito, tipo: 'debito' };
    if (vCredito > 0) return { valor: vCredito, tipo: 'credito' };
    return { valor: 0, tipo: null };
  }

  // Coluna única "valor": usa o sinal (quando existir) para inferir o tipo
  const bruto = parseValorComSinal(encontrarValor(linha, [
    'valor', 'value', 'amount', 'vlr', 'valor (r$)', 'valor r$'
  ]));
  if (bruto === 0) return { valor: 0, tipo: null };
  return { valor: Math.abs(bruto), tipo: bruto < 0 ? 'debito' : 'credito' };
}

function formatarData(str) {
  if (!str) return '';
  str = String(str).trim();

  // Excel serial date
  if (!isNaN(str) && Number(str) > 30000) {
    try {
      const excelDate = XLSX.SSF.parse_date_code(Number(str));
      if (excelDate) {
        return `${excelDate.y}-${String(excelDate.m).padStart(2,'0')}-${String(excelDate.d).padStart(2,'0')}`;
      }
    } catch(e) {}
  }

  if (str.includes('/')) {
    const partes = str.split('/');
    if (partes.length === 3) {
      let [d, m, a] = partes;
      if (a.length === 2) a = '20' + a;
      return `${a}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }
  }
  return str;
}

// Escapa texto vindo dos arquivos do usuário antes de jogar no innerHTML,
// evitando que um nome de fornecedor com <, >, & ou aspas quebre a tabela.
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatarMoeda(valor) {
  return Number(valor).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatarDataBR(dataISO) {
  if (!dataISO || dataISO.length < 8) return '-';
  const partes = dataISO.split('-');
  if (partes.length !== 3) return dataISO;
  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

// -------------------- MATCHING --------------------
function conciliar() {
  console.log('Iniciando conciliação...');
  console.log('Extrato:', estado.extrato.length, '| Fornecedores:', estado.fornecedores.length);

  try {
    estado.extrato.forEach(i => i.usado = false);
    estado.fornecedores.forEach(i => i.usado = false);

    const conciliados = [];
    const pendentes = [];

    // 1. Match forte: documento + valor
    for (const forn of estado.fornecedores) {
      if (forn.usado) continue;

      const match = estado.extrato.find(ext =>
        !ext.usado &&
        ext.documento &&
        forn.documento &&
        ext.documento === forn.documento &&
        Math.abs(ext.valor - forn.valor) < 0.02
      );

      if (match) {
        match.usado = true;
        forn.usado = true;
        conciliados.push(criarMatch(match, forn, 'documento+valor'));
      }
    }

    // 2. Match por valor + nome — agora GLOBAL: monta todos os pares candidatos
    // válidos, ordena do melhor score para o pior e atribui nessa ordem. Isso evita
    // que o primeiro fornecedor da planilha "roube" um match que serviria melhor
    // para outro fornecedor mais abaixo, quando há valores repetidos.
    const candidatos = [];
    for (const forn of estado.fornecedores) {
      if (forn.usado) continue;
      for (const ext of estado.extrato) {
        if (ext.usado) continue;
        if (Math.abs(ext.valor - forn.valor) > 0.02) continue;
        // Se soubermos o tipo da movimentação, só casa com saídas (pagamentos),
        // nunca com entradas/créditos no extrato.
        if (ext.tipo === 'credito') continue;

        const scoreNome = similaridade(ext.descricao, forn.fornecedor);
        if (scoreNome < 0.35) continue;

        const dias = diferencaDias(ext.data, forn.data);
        // pequeno bônus de desempate quando as datas estão próximas
        const score = scoreNome + (dias !== null ? Math.max(0, (5 - dias)) / 100 : 0);

        candidatos.push({ forn, ext, score });
      }
    }

    candidatos.sort((a, b) => b.score - a.score);

    for (const cand of candidatos) {
      if (cand.forn.usado || cand.ext.usado) continue;
      cand.ext.usado = true;
      cand.forn.usado = true;
      conciliados.push(criarMatch(cand.ext, cand.forn, 'valor+nome'));
    }

    // 3. Pendentes
    estado.extrato.filter(i => !i.usado).forEach(i => {
      pendentes.push({ ...i, tipo: 'extrato' });
    });
    estado.fornecedores.filter(i => !i.usado).forEach(i => {
      pendentes.push({ ...i, tipo: 'fornecedor' });
    });

    estado.conciliados = conciliados;
    estado.pendentes = pendentes;
    estado.lancamentos = gerarLancamentos(conciliados);

    console.log('Resultado → Conciliados:', conciliados.length, '| Pendentes:', pendentes.length);

    atualizarResumo();
    renderizarTabelas();
    mostrarTela(telaResultado);

  } catch (err) {
    console.error('Erro na conciliação:', err);
    alert('Erro durante a conciliação. Veja o Console (F12).');
  }
}

function diferencaDias(dataA, dataB) {
  if (!dataA || !dataB) return null;
  const a = new Date(dataA);
  const b = new Date(dataB);
  if (isNaN(a) || isNaN(b)) return null;
  return Math.abs((a - b) / (1000 * 60 * 60 * 24));
}

function criarMatch(ext, forn, tipo) {
  return {
    id: `match-${ext.id}-${forn.id}`,
    data: forn.data || ext.data,
    documento: forn.documento || ext.documento,
    fornecedor: forn.fornecedor,
    descricao: ext.descricao,
    valor: forn.valor,
    tipoMatch: tipo,
    extratoId: ext.id,
    fornecedorId: forn.id
  };
}

function similaridade(str1, str2) {
  if (!str1 || !str2) return 0;
  const a = str1.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const b = str2.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  if (a.includes(b) || b.includes(a)) return 0.9;

  const palavrasA = a.split(/\s+/).filter(p => p.length > 2);
  const palavrasB = b.split(/\s+/).filter(p => p.length > 2);
  if (palavrasA.length === 0 || palavrasB.length === 0) return 0;

  let comuns = 0;
  palavrasA.forEach(p => {
    if (palavrasB.some(pb => pb.includes(p) || p.includes(pb))) comuns++;
  });
  return comuns / Math.max(palavrasA.length, palavrasB.length);
}

function gerarLancamentos(conciliados) {
  return conciliados.map(item => ({
    id: `lanc-${item.id}`,
    data: item.data,
    debito: '148',
    participante: item.fornecedor,
    credito: 'Banco',
    valor: item.valor,
    origem: item.tipoMatch === 'manual' ? 'manual' : 'automático'
  }));
}

// -------------------- BOTÃO CONCILIAR --------------------
btnConciliar.addEventListener('click', () => {
  if (!estado.extrato.length || !estado.fornecedores.length) {
    alert('Envie os dois arquivos antes de conciliar.');
    return;
  }
  btnConciliar.disabled = true;
  btnConciliar.textContent = 'Conciliando...';

  setTimeout(() => {
    conciliar();
    btnConciliar.disabled = false;
    btnConciliar.textContent = 'Conciliar Automaticamente';
  }, 80);
});

// -------------------- RENDERIZAÇÃO --------------------
function atualizarResumo() {
  $('#qtd-conciliados').textContent = estado.conciliados.length;
  $('#qtd-pendentes').textContent = estado.pendentes.length;
  const total = estado.conciliados.reduce((acc, i) => acc + i.valor, 0);
  $('#valor-total').textContent = formatarMoeda(total);
}

function renderizarTabelas() {
  // Conciliados
  const tbodyConc = $('#tbody-conciliados');
  if (estado.conciliados.length === 0) {
    tbodyConc.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#6b7280;padding:2rem;">Nenhum item conciliado</td></tr>`;
  } else {
    tbodyConc.innerHTML = estado.conciliados.map(item => `
      <tr>
        <td>${formatarDataBR(item.data)}</td>
        <td>${escapeHtml(item.documento) || '-'}</td>
        <td>${escapeHtml(item.fornecedor) || '-'}</td>
        <td>${formatarMoeda(item.valor)}</td>
        <td><span class="badge badge-success">Conciliado</span></td>
        <td><button class="btn btn-outline btn-sm" onclick="abrirModalCorrecao('${item.id}')">Corrigir</button></td>
      </tr>
    `).join('');
  }

  // Pendentes
  const tbodyPend = $('#tbody-pendentes');
  if (estado.pendentes.length === 0) {
    tbodyPend.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#6b7280;padding:2rem;">Nenhum item pendente</td></tr>`;
  } else {
    tbodyPend.innerHTML = estado.pendentes.map(item => `
      <tr>
        <td><span class="badge ${item.tipo === 'extrato' ? 'badge-info' : 'badge-warning'}">${item.tipo === 'extrato' ? 'Extrato' : 'Fornecedor'}</span></td>
        <td>${formatarDataBR(item.data)}</td>
        <td>${escapeHtml(item.documento) || '-'}</td>
        <td>${escapeHtml(item.fornecedor || item.descricao) || '-'}</td>
        <td>${formatarMoeda(item.valor)}</td>
        <td><button class="btn btn-outline btn-sm" onclick="forcarConciliacao('${item.id}')">Forçar</button></td>
      </tr>
    `).join('');
  }

  // Lançamentos
  const tbodyLanc = $('#tbody-lancamentos');
  if (estado.lancamentos.length === 0) {
    tbodyLanc.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#6b7280;padding:2rem;">Nenhum lançamento gerado</td></tr>`;
  } else {
    tbodyLanc.innerHTML = estado.lancamentos.map(item => `
      <tr>
        <td>${formatarDataBR(item.data)}</td>
        <td>${item.debito}</td>
        <td>${escapeHtml(item.participante) || '-'}</td>
        <td>${item.credito}</td>
        <td>${formatarMoeda(item.valor)}</td>
        <td><span class="badge ${item.origem === 'manual' ? 'badge-warning' : 'badge-info'}">${item.origem}</span></td>
      </tr>
    `).join('');
  }
}

// -------------------- TABS --------------------
$$('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    $$('.tab').forEach(t => t.classList.remove('ativa'));
    $$('.tab-pane').forEach(p => p.classList.remove('ativa'));
    tab.classList.add('ativa');
    $(`#tab-${tab.dataset.tab}`).classList.add('ativa');
  });
});

// -------------------- MODAL --------------------
function abrirModalCorrecao(id) {
  const item = estado.conciliados.find(i => i.id === id);
  if (!item) return;

  estado.itemEditando = item;
  $('#modal-documento').value = item.documento || '';
  $('#modal-fornecedor').value = item.fornecedor || '';
  $('#modal-valor').value = item.valor.toFixed(2).replace('.', ',');
  $('#modal-data').value = item.data || '';
  $('#modal-correcao').classList.remove('hidden');
}

function fecharModal() {
  $('#modal-correcao').classList.add('hidden');
  estado.itemEditando = null;
}

$('#btn-fechar-modal').addEventListener('click', fecharModal);
$('#btn-cancelar-modal').addEventListener('click', fecharModal);
$('#modal-correcao .modal-overlay').addEventListener('click', fecharModal);

$('#btn-salvar-correcao').addEventListener('click', () => {
  if (!estado.itemEditando) return;

  const item = estado.itemEditando;
  item.documento = $('#modal-documento').value;
  item.fornecedor = $('#modal-fornecedor').value;
  item.valor = parseValor($('#modal-valor').value);
  item.data = $('#modal-data').value;
  item.tipoMatch = 'manual';

  const lanc = estado.lancamentos.find(l => l.id === `lanc-${item.id}`);
  if (lanc) {
    lanc.participante = item.fornecedor;
    lanc.valor = item.valor;
    lanc.data = item.data;
    lanc.origem = 'manual';
  }

  atualizarResumo();
  renderizarTabelas();
  fecharModal();
});

// -------------------- MODAL: FORÇAR CONCILIAÇÃO MANUAL --------------------
let forcarItemOrigem = null;

function resumoItemPendente(item) {
  if (item.tipo === 'extrato') {
    return `Extrato — ${formatarDataBR(item.data)} — ${item.descricao || 'sem descrição'} — ${formatarMoeda(item.valor)}`;
  }
  return `Fornecedor — ${formatarDataBR(item.data)} — ${item.fornecedor || 'sem nome'} — ${formatarMoeda(item.valor)}`;
}

function forcarConciliacao(id) {
  const item = estado.pendentes.find(i => i.id === id);
  if (!item) return;

  const ehExtrato = item.tipo === 'extrato';
  const listaOposta = estado.pendentes.filter(p => p.tipo === (ehExtrato ? 'fornecedor' : 'extrato'));

  if (!listaOposta.length) {
    alert(`Não há itens pendentes do lado ${ehExtrato ? 'de fornecedores' : 'do extrato'} para vincular.`);
    return;
  }

  forcarItemOrigem = item;

  // textContent evita qualquer risco de HTML vindo do arquivo do usuário
  $('#forcar-item-resumo').textContent = resumoItemPendente(item);

  const select = $('#forcar-select');
  select.innerHTML = '';
  listaOposta
    .slice()
    .sort((a, b) => Math.abs(a.valor - item.valor) - Math.abs(b.valor - item.valor))
    .forEach(opp => {
      const opt = document.createElement('option');
      opt.value = opp.id;
      opt.textContent = resumoItemPendente(opp);
      select.appendChild(opt);
    });

  $('#btn-confirmar-forcar').disabled = true;
  $('#modal-forcar').classList.remove('hidden');
}

function fecharModalForcar() {
  $('#modal-forcar').classList.add('hidden');
  $('#forcar-select').innerHTML = '';
  $('#btn-confirmar-forcar').disabled = true;
  forcarItemOrigem = null;
}

$('#forcar-select').addEventListener('change', () => {
  $('#btn-confirmar-forcar').disabled = !$('#forcar-select').value;
});

$('#btn-fechar-modal-forcar').addEventListener('click', fecharModalForcar);
$('#btn-cancelar-forcar').addEventListener('click', fecharModalForcar);
$('#modal-forcar .modal-overlay').addEventListener('click', fecharModalForcar);

$('#btn-confirmar-forcar').addEventListener('click', () => {
  const outroId = $('#forcar-select').value;
  if (!outroId || !forcarItemOrigem) return;

  const outro = estado.pendentes.find(i => i.id === outroId);
  if (!outro) return;

  const ext = forcarItemOrigem.tipo === 'extrato' ? forcarItemOrigem : outro;
  const forn = forcarItemOrigem.tipo === 'fornecedor' ? forcarItemOrigem : outro;

  const match = criarMatch(ext, forn, 'manual');
  estado.conciliados.push(match);
  estado.lancamentos.push(...gerarLancamentos([match]));

  const idsRemover = new Set([forcarItemOrigem.id, outro.id]);
  estado.pendentes = estado.pendentes.filter(p => !idsRemover.has(p.id));

  atualizarResumo();
  renderizarTabelas();
  fecharModalForcar();
});

// -------------------- HISTÓRICO + EXPORTAR --------------------
function despojarRaw(item) {
  // remove o campo "raw" (linha original do arquivo) antes de persistir,
  // para não estourar o limite do localStorage
  const { raw, ...resto } = item;
  return resto;
}

$('#btn-salvar').addEventListener('click', () => {
  const registro = {
    id: Date.now(),
    data: new Date().toLocaleString('pt-BR'),
    conciliados: estado.conciliados.length,
    pendentes: estado.pendentes.length,
    valorTotal: estado.conciliados.reduce((acc, i) => acc + i.valor, 0),
    dados: {
      conciliados: estado.conciliados.map(despojarRaw),
      pendentes: estado.pendentes.map(despojarRaw),
      lancamentos: estado.lancamentos
    }
  };
  estado.historico.unshift(registro);
  // mantém só os últimos 30 registros para não estourar o localStorage
  estado.historico = estado.historico.slice(0, 30);

  try {
    localStorage.setItem('conciliacao_historico', JSON.stringify(estado.historico));
    alert('Histórico salvo com sucesso!');
  } catch (err) {
    console.error('Erro ao salvar histórico:', err);
    alert('Não foi possível salvar o histórico — o armazenamento local pode estar cheio. Tente excluir conciliações antigas.');
  }
});

function renderizarHistorico() {
  const container = $('#lista-historico');
  if (!estado.historico.length) {
    container.innerHTML = `<div class="empty-state"><p>Nenhuma conciliação salva ainda.</p></div>`;
    return;
  }
  container.innerHTML = estado.historico.map(h => `
    <div class="historico-item">
      <div onclick="carregarHistorico(${h.id})" style="cursor:pointer;flex:1;">
        <strong>${h.data}</strong>
        <div style="font-size:0.85rem;color:#6b7280;margin-top:0.2rem;">
          ${h.conciliados} conciliados • ${h.pendentes} pendentes
        </div>
      </div>
      <div style="font-weight:600;color:#2563eb;margin-right:1rem;">${formatarMoeda(h.valorTotal)}</div>
      <button class="btn-remove" title="Excluir do histórico" onclick="excluirHistorico(event, ${h.id})">✕</button>
    </div>
  `).join('');
}

function excluirHistorico(event, id) {
  event.stopPropagation();
  if (!confirm('Excluir esta conciliação do histórico?')) return;
  estado.historico = estado.historico.filter(h => h.id !== id);
  localStorage.setItem('conciliacao_historico', JSON.stringify(estado.historico));
  renderizarHistorico();
}

function carregarHistorico(id) {
  const registro = estado.historico.find(h => h.id === id);
  if (!registro) return;
  estado.conciliados = registro.dados.conciliados;
  estado.pendentes = registro.dados.pendentes;
  estado.lancamentos = registro.dados.lancamentos;
  atualizarResumo();
  renderizarTabelas();
  mostrarTela(telaResultado);
}

$('#btn-exportar').addEventListener('click', () => {
  if (!estado.lancamentos.length) {
    alert('Nenhum lançamento para exportar.');
    return;
  }
  let csv = 'Data;Débito;Participante;Crédito;Valor;Origem\n';
  estado.lancamentos.forEach(l => {
    csv += `${formatarDataBR(l.data)};${l.debito};${l.participante};${l.credito};${l.valor.toFixed(2).replace('.', ',')};${l.origem}\n`;
  });
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `lancamentos_${new Date().toISOString().slice(0,10)}.csv`;
  link.click();
});

// Expor funções globais
window.abrirModalCorrecao = abrirModalCorrecao;
window.forcarConciliacao = forcarConciliacao;
window.carregarHistorico = carregarHistorico;
window.excluirHistorico = excluirHistorico;