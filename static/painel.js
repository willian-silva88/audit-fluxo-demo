'use strict';

const $ = (id) => document.getElementById(id);
const moeda = new Intl.NumberFormat('pt-BR', {style: 'currency', currency: 'BRL'});
const valor = (centavos) => moeda.format(centavos / 100);
const rotulos = {
  encontrado: 'Baixa encontrada',
  conferir: 'Requer conferência',
  nao_encontrado: 'Sem correspondência',
  revisado: 'Vínculo conferido',
};
let estado;
let atual;
let ocupado = false;

function no(tag, texto = '', classe = '') {
  const el = document.createElement(tag);
  el.textContent = texto;
  el.className = classe;
  return el;
}

function botao(texto, callback, classe = '') {
  const el = no('button', texto, classe);
  el.type = 'button';
  el.addEventListener('click', callback);
  return el;
}

function situacao(item) {
  return item.review_status === 'conferido' ? 'revisado' : item.match_status;
}

function selo(item) {
  const status = situacao(item);
  return no('span', rotulos[status], 'estado ' + status);
}

function mensagem(texto, erro = false) {
  const el = $('mensagem');
  el.hidden = !texto;
  el.textContent = texto;
  el.classList.toggle('erro', erro);
}

async function api(url, dados) {
  const resposta = await fetch(url, dados === undefined ? undefined : {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(dados),
  });
  const json = await resposta.json();
  if (!resposta.ok) throw new Error(json.erro);
  return json;
}

async function acao(callback) {
  if (ocupado) return;
  ocupado = true;
  $('carregar').disabled = true;
  mensagem('');
  try {
    await callback();
  } catch (erro) {
    mensagem(erro.message, true);
  } finally {
    ocupado = false;
    $('carregar').disabled = false;
  }
}

function renderPainel() {
  const itens = estado.requests;
  $('painel').hidden = $('exportar').disabled = !itens.length;
  $('vazio').hidden = !!itens.length;
  const indicadores = [
    ['Documentos solicitados', itens.length],
    ['Requerem conferência', itens.filter((item) => situacao(item) === 'conferir').length],
    ['Sem correspondência', estado.stats.not_found],
    ['Vínculos conferidos', estado.stats.reviewed],
  ];
  $('indicadores').replaceChildren(...indicadores.map(([label, quantidade]) => {
    const card = no('div', '', 'indicador');
    card.append(no('span', label), no('strong', quantidade));
    return card;
  }));
  const filtro = $('filtro').value;
  const filtrados = itens.filter((item) => !filtro || filtro === situacao(item));
  if (!filtrados.length) {
    $('documentos').replaceChildren(no('p', 'Nenhum documento neste filtro.', 'cabecalho-cartao'));
    return;
  }
  $('documentos').replaceChildren(...filtrados.map((item) => {
    const linha = botao('', () => acao(() => abrir(item.id)), 'documento');
    linha.setAttribute('aria-pressed', String(atual?.id === item.id));
    const texto = no('span', item.invoice);
    texto.append(no('small', item.supplier));
    linha.append(texto, selo(item));
    return linha;
  }));
}

function blocoValor(label, quantia) {
  const bloco = no('div');
  bloco.append(no('dt', label), no('dd', valor(quantia)));
  return bloco;
}

function tabela(item, checagens) {
  const resultado = no('table');
  const topo = resultado.createTHead().insertRow();
  for (const texto of ['Vincular', 'Filial / tipo', 'Data da baixa', 'Banco / conta da base', 'Líquido']) {
    const th = no('th', texto);
    th.scope = 'col';
    topo.append(th);
  }
  const corpo = resultado.createTBody();
  for (const baixa of item.matches) {
    const linha = corpo.insertRow();
    const check = no('input');
    check.type = 'checkbox';
    check.disabled = !baixa.eligible || !item.branch;
    check.checked = baixa.selected;
    check.setAttribute('aria-label', `Vincular ${baixa.title} da filial ${baixa.branch}`);
    checagens.set(baixa, check);
    linha.insertCell().append(check);
    linha.insertCell().textContent = `${baixa.branch} / ${baixa.document_type || '—'}`;
    linha.insertCell().textContent = baixa.payment_date?.split('-').reverse().join('/') || '—';
    const conta = linha.insertCell();
    conta.textContent = baixa.account || 'Não informada';
    if (!baixa.eligible) conta.append(no('small', 'Tipo fora do critério'));
    else if (baixa.exchange) conta.append(no('small', 'Permuta / adiantamento'));
    linha.insertCell().textContent = valor(baixa.amount_cents);
  }
  const rolagem = no('div', '', 'rolagem');
  rolagem.append(resultado);
  return rolagem;
}

function renderDetalhe() {
  const detalhe = $('detalhe');
  const item = atual;
  if (!item) {
    detalhe.replaceChildren(no('p', 'Selecione um documento para analisar as baixas.'));
    return;
  }
  const topo = no('div', '', 'detalhe-topo');
  topo.append(no('h2', `${item.invoice} · ${item.supplier}`), selo(item));
  const valores = no('dl', '', 'valores');
  const selecionado = blocoValor('Líquido selecionado', item.selected_amount_cents);
  valores.append(blocoValor('Valor da nota', item.amount_cents),
    blocoValor('Líquido localizado', item.matched_amount_cents), selecionado);
  detalhe.replaceChildren(no('p', '02 / CONFERÊNCIA', 'rotulo'), topo,
    no('p', `${item.code} · Filial: ${item.branch || 'não informada'} · Tipo aceito: ${estado.config.document_type}`, 'subtitulo'),
    valores);
  if (item.warnings.length) {
    const alertas = no('ul', '', 'alertas');
    alertas.append(...item.warnings.map((aviso) => no('li', aviso.label)));
    detalhe.append(alertas);
  }
  const checagens = new Map();
  if (!item.matches.length) {
    detalhe.append(no('p', 'A base fictícia não tem candidatos para este código e filial.'));
    return;
  }
  detalhe.append(no('h3', 'Baixas candidatas'), tabela(item, checagens));
  if (!item.branch) {
    const filiais = [...new Set(item.matches.filter((baixa) => baixa.eligible).map((baixa) => baixa.branch))];
    if (!filiais.length) return;
    const barra = no('div', '', 'form-linha');
    const label = no('label', 'Filial confirmada para o documento: ');
    const campo = no('select');
    const escolher = no('option', 'Selecione…');
    escolher.value = '';
    campo.append(escolher, ...filiais.map((filial) => no('option', filial)));
    const salvar = botao('Aplicar filial', () => acao(() => atualizar({
      id: item.id, branch: campo.value,
    }, 'Filial registrada. Revise as baixas restantes.')));
    salvar.disabled = true;
    campo.addEventListener('change', () => { salvar.disabled = !campo.value; });
    label.append(campo);
    barra.append(label, salvar);
    detalhe.append(barra);
    return;
  }
  if (!item.matches_count) return;
  const barra = no('div', '', 'form-linha');
  const label = no('label', 'Justificativa (obrigatória para exceções)');
  const campo = no('textarea');
  campo.rows = 1;
  campo.maxLength = 500;
  campo.value = item.justification || '';
  label.append(campo);
  const confirmar = botao('Confirmar vínculo', () => acao(() => atualizar({
    id: item.id,
    justification: campo.value,
    selected_match_ids: [...checagens].filter(([, check]) => check.checked).map(([baixa]) => baixa.id),
    review_confirmed: true,
  }, 'Vínculo conferido. A etapa de localizar o comprovante continua separada.')), 'primario');
  const reverTotal = () => {
    const marcadas = [...checagens].filter(([, check]) => check.checked);
    selecionado.lastChild.textContent = valor(marcadas.reduce((soma, [baixa]) => soma + baixa.amount_cents, 0));
    confirmar.disabled = !marcadas.length;
  };
  for (const check of checagens.values()) check.addEventListener('change', reverTotal);
  reverTotal();
  barra.append(label, confirmar);
  detalhe.append(barra);
  if (item.review_status === 'conferido') detalhe.append(no('p', '✓ Decisão registrada no histórico desta sessão.', 'historico'));
}

async function abrir(id) {
  atual = id ? await api('/api/documento/' + id) : null;
  renderPainel();
  renderDetalhe();
}

async function abrirFiltro(preferido) {
  const filtro = $('filtro').value;
  const visiveis = estado.requests.filter((item) => !filtro || filtro === situacao(item));
  await abrir(visiveis.find((item) => item.id === preferido)?.id || visiveis[0]?.id);
}

async function atualizar(payload, texto) {
  atual = await api('/api/conferencia', payload);
  estado = await api('/api/painel');
  if ($('filtro').value && $('filtro').value !== situacao(atual)) $('filtro').value = '';
  renderPainel();
  renderDetalhe();
  mensagem(texto);
}

$('carregar').addEventListener('click', () => acao(async () => {
  estado = await api('/api/exemplo', {});
  $('filtro').value = '';
  await abrir(estado.requests.find((item) => item.match_status === 'conferir')?.id);
  $('carregar').textContent = 'Recarregar exemplo';
}));
$('filtro').addEventListener('change', () => acao(() => abrirFiltro(atual?.id)));
$('exportar').addEventListener('click', () => { window.location.assign('/api/planilha'); });
acao(async () => {
  estado = await api('/api/painel');
  if (estado.requests.length) {
    $('carregar').textContent = 'Recarregar exemplo';
    await abrirFiltro();
  }
});
