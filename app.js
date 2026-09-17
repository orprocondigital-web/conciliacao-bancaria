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
    
    const valor = parseValor(encontrarValor(linha, [
      'valor', 'value', 'amount', 'vlr', 'valor (r$)', 'valor r$', 
      'crédito', 'débito', 'credito', 'debito', 'entrada', 'saída'
    ]));

    return {
      id: `ext-${idx}`,
      origem: 'extrato',
      data: formatarData(data),
      documento: String(documento || '').replace(/\D/g, ''),
      descricao: descricao || '',
      valor: valor,
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

function parseValor(str) {
  if (str === null || str === undefined || str === '') return 0;
  let limpo = String(str)
    .replace(/R\$\s?/gi, '')
    .replace(/\s/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const num = parseFloat(limpo);
  return isNaN(num) ? 0 : Math.abs(num);
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

    // 2. Match por valor + nome
    for (const forn of estado.fornecedores) {
      if (forn.usado) continue;

      let melhor = null;
      let melhorScore = 0;

      for (const ext of estado.extrato) {
        if (ext.usado) continue;
        if (Math.abs(ext.valor - forn.valor) > 0.02) continue;

        const score = similaridade(ext.descricao, forn.fornecedor);
        if (score > melhorScore && score >= 0.35) {
          melhorScore = score;
          melhor = ext;
        }
      }

      if (melhor) {
        melhor.usado = true;
        forn.usado = true;
        conciliados.push(criarMatch(melhor, forn, 'valor+nome'));
      }
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
        <td>${item.documento || '-'}</td>
        <td>${item.fornecedor || '-'}</td>
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
        <td>${item.documento || '-'}</td>
        <td>${item.fornecedor || item.descricao || '-'}</td>
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
        <td>${item.participante || '-'}</td>
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
$('.modal-overlay').addEventListener('click', fecharModal);

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

function forcarConciliacao(id) {
  alert('Em breve: funcionalidade de forçar match manual.');
}

// -------------------- HISTÓRICO + EXPORTAR --------------------
$('#btn-salvar').addEventListener('click', () => {
  const registro = {
    id: Date.now(),
    data: new Date().toLocaleString('pt-BR'),
    conciliados: estado.conciliados.length,
    pendentes: estado.pendentes.length,
    valorTotal: estado.conciliados.reduce((acc, i) => acc + i.valor, 0),
    dados: {
      conciliados: estado.conciliados,
      pendentes: estado.pendentes,
      lancamentos: estado.lancamentos
    }
  };
  estado.historico.unshift(registro);
  localStorage.setItem('conciliacao_historico', JSON.stringify(estado.historico));
  alert('Histórico salvo com sucesso!');
});

function renderizarHistorico() {
  const container = $('#lista-historico');
  if (!estado.historico.length) {
    container.innerHTML = `<div class="empty-state"><p>Nenhuma conciliação salva ainda.</p></div>`;
    return;
  }
  container.innerHTML = estado.historico.map(h => `
    <div class="historico-item" onclick="carregarHistorico(${h.id})">
      <div>
        <strong>${h.data}</strong>
        <div style="font-size:0.85rem;color:#6b7280;margin-top:0.2rem;">
          ${h.conciliados} conciliados • ${h.pendentes} pendentes
        </div>
      </div>
      <div style="font-weight:600;color:#2563eb;">${formatarMoeda(h.valorTotal)}</div>
    </div>
  `).join('');
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

// Estilo extra
const styleExtra = document.createElement('style');
styleExtra.textContent = `.btn-sm { padding: 0.3rem 0.7rem; font-size: 0.8rem; }`;
document.head.appendChild(styleExtra);

// Expor funções globais
window.abrirModalCorrecao = abrirModalCorrecao;
window.forcarConciliacao = forcarConciliacao;
window.carregarHistorico = carregarHistorico;