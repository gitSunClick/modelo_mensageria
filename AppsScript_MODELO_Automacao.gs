/**
 * SunClick - MODELO DE AUTOMACAO: COLAR relatorio > ORGANIZAR > ENVIAR para o Bitrix24
 * (baseado na Regua de Cobranca v2; a mensagem de WhatsApp e enviada pelo robo do Bitrix quando o card entra na etapa)
 *
 * COMO USAR ESTE MODELO
 *  1. Preencha so o bloco "CONFIGURACAO" logo abaixo (webhook, funil, responsavel, etapas).
 *  2. Salve, recarregue a planilha e rode o menu "Testar configuracao". Ele confere webhook, funil e etapas.
 *  3. Rode "0. Preparar aba Colar aqui" (1x) e teste com uma planilha pequena com o SEU telefone.
 *  O que NAO precisa mexer: leitura de colunas, limpeza, chave anti-duplicidade e criacao dos campos no Bitrix.
 *  Se a sua regra de etapas for diferente (ex.: uma etapa so), ajuste apenas REGUA e etapaPorDias.
 *
 * Aba "Colar aqui": cola o relatorio do jeito que sai do sistema (com ou sem cabecalho).
 * Botao "1. Organizar": reconhece as colunas pelo CONTEUDO (telefone, data, valor, nome) e monta a aba
 *   "Cobranca" (A NOME | B TELEFONE | C VALOR | D VENCIMENTO | E ETAPA | F STATUS).
 * Botao "2. Enviar": cria contato + card na etapa certa, sem duplicar.
 */

// ===================== CONFIGURACAO (preencha aqui) =====================
const WEBHOOK_URL = 'COLE_AQUI_O_WEBHOOK_DE_ENTRADA_DO_BITRIX/';   // Bitrix > Recursos para desenvolvedores > Webhook de entrada (so CRM). Termina com "/"
const CATEGORY_ID = 'COLE_AQUI_O_ID_DO_FUNIL';                     // numero depois de "category/" no endereco do funil (ex.: '38')
const RESPONSAVEL_ID = 'COLE_AQUI_O_ID_DO_USUARIO';                // numero depois de "user/" no perfil da pessoa (ex.: '3756')
// dias de atraso desde o vencimento -> inicio do nome da etapa no Bitrix (a etapa pode se chamar D+5 ou D5)
const REGUA = { '0': 'D0', '5': 'D+5', '10': 'D+10', '15': 'D+15', '30': 'D+30' };
const AGRUPAR_POR_CLIENTE = true;   // true: 1 mensagem por cliente por dia (etapa mais avancada)
// Nomes das abas (pode manter)
const ABA_COLAR = 'Colar aqui';
const ABA_COBRANCA = 'Cobranca';
// ========================================================================

// Atraso de 30 dias ou mais cai sempre na ultima etapa (D+30)
function etapaPorDias(d) { return d >= 30 ? REGUA['30'] : REGUA[String(d)]; }

const CAMPO_CHAVE = 'UF_CRM_CHAVE_FATURA';   // campo "Chave da fatura" no negocio: evita card duplicado (criado sozinho no primeiro envio)
const CAMPO_ETAPA = 'UF_CRM_ETAPA_REGUA';   // campo "Etapa da regua" no negocio (criado sozinho no primeiro envio)

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡ SunClick ⚡')
    .addItem('📋 0. Preparar aba Colar aqui (1x)', 'prepararAbaColar')
    .addItem('🧹 1. Organizar o que colei', 'organizar')
    .addItem('🚀 2. Enviar Registros', 'enviarParaBitrix')
    .addSeparator()
    .addItem('🔧 Testar configuração', 'testarConfiguracao')
    .addToUi();
}

function enviarParaBitrix() {
  const ui = SpreadsheetApp.getUi();
  const folha = SpreadsheetApp.getActive().getSheetByName(ABA_COBRANCA);
  if (!folha) { ui.alert('Antes, use "1. Organizar o que colei".'); return; }
  const ultimaLinha = folha.getLastRow();
  if (ultimaLinha < 2) { ui.alert('Não há dados para processar.'); return; }

  const falhou = garantirCampos();   // cria sozinho os campos que faltam no Bitrix (Vencimento, Etapa, Chave)
  if (falhou.length) {
    ui.alert('Nao consegui criar o(s) campo(s) no Bitrix: ' + falhou.join(', ') + '.\n\nConfira se o webhook tem permissao de CRM.');
    return;
  }
  const dados = folha.getRange(2, 1, ultimaLinha - 1, 6).getValues();
  const etapas = buscarEtapas();
  const hoje = hojeUTC();

  // 1) classifica cada linha e conta por etapa
  const fila = [];
  const contagem = {};
  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    if (!linha[0]) continue;                               // sem nome: ignora
    if (String(linha[5]).indexOf('OK') === 0) continue;    // já enviado
    if (String(linha[5]).indexOf('AGRUPADO') === 0) continue;   // cliente recebe so a etapa mais avancada
    const ven = lerData(linha[3]);
    const dias = ven ? Math.round((hoje - ven.utc) / 86400000) : null;
    const etapa = dias === null ? null : etapaPorDias(dias);
    if (!etapa) {
      folha.getRange(i + 2, 5).setValue('');
      folha.getRange(i + 2, 6).setValue(ven ? 'FORA DA REGUA (' + dias + ' dias)' : 'ERRO: data de vencimento invalida');
      continue;
    }
    const stageId = acharEtapa(etapas, etapa);
    if (!stageId) {
      folha.getRange(i + 2, 6).setValue('ERRO: etapa ' + etapa + ' nao existe no funil');
      continue;
    }
    folha.getRange(i + 2, 5).setValue(tagEtapa(etapa));   // tag: D0, D5, D10, D15, D30
    fila.push({ linha: i + 2, dados: linha, venc: ven.texto, etapa: etapa, stageId: stageId });
    contagem[etapa] = (contagem[etapa] || 0) + 1;
  }
  if (!fila.length) { ui.alert('Nada para enviar. Veja a coluna F.'); return; }

  // 2) confirma com o resumo
  const resumo = Object.keys(contagem).map(function (e) { return e + ': ' + contagem[e]; }).join(' | ');
  const resp = ui.alert('Confirmar envio', fila.length + ' registro(s)\n' + resumo +
    '\n\nOs robôs vão disparar o WhatsApp agora. Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // 3) envia
  let enviados = 0, jaExistiam = 0;
  for (let k = 0; k < fila.length; k++) {
    const item = fila[k];
    try {
      const fone = limparFone(item.dados[1]);
      if (!/^55\d{10,11}$/.test(fone)) throw new Error('telefone invalido');
      const valor = String(item.dados[2]).replace(/R\$\s*/gi, '').trim();
      if (!valor) throw new Error('valor vazio');
      // 1) esta fatura ja tem card aberto no funil? (chave = telefone + valor + vencimento)
      const chave = chaveFatura(fone, valor, item.venc);
      const ex = buscarNegocioPorChave(chave);
      if (ex) {
        const msg = tratarNegocioExistente(ex, item);
        folha.getRange(item.linha, 6).setValue(msg);
        if (/movido/.test(msg)) enviados++; else jaExistiam++;
      } else {
        // 2) nao tem: reaproveita o contato (pelo telefone) e cria o card
        const contatoId = criarContato(item.dados[0], fone);
        const negocioId = criarNegocio(item, contatoId, valor, chave);
        folha.getRange(item.linha, 6).setValue('OK negocio ' + negocioId);
        enviados++;
      }
    } catch (e) {
      folha.getRange(item.linha, 6).setValue('ERRO: ' + e.message);
    }
    Utilities.sleep(1000);   // anti-flood
  }
  ui.alert('Processamento Concluído!\n\n' + enviados + ' registro(s) enviado(s) com sucesso (card novo ou movido de etapa).' +
    (jaExistiam ? '\n' + jaExistiam + ' ja existia(m) no funil e nao foi(ram) duplicado(s).' : ''));
}

// --- Funções de API (Contact & Deal) ---

function criarContato(nome, fone) {
  // se o telefone já existe no Bitrix, reaproveita o contato
  const dup = chamar('crm.duplicate.findbycomm', { type: 'PHONE', values: [fone] });
  if (dup.result && dup.result.CONTACT && dup.result.CONTACT.length) return dup.result.CONTACT[0];

  const r = chamar('crm.contact.add', { fields: {
    NAME: String(nome).trim(), OPENED: 'Y', ASSIGNED_BY_ID: RESPONSAVEL_ID,
    PHONE: [{ VALUE: fone, VALUE_TYPE: 'WORK' }]
  } });
  if (!r.result) throw new Error('contato: ' + JSON.stringify(r));
  return r.result;
}

function criarNegocio(item, contatoId, valor, chave) {
  const r = chamar('crm.deal.add', { fields: {
    TITLE: 'Cobranca ' + item.etapa + ' - ' + String(item.dados[0]).trim(),
    CATEGORY_ID: CATEGORY_ID,
    STAGE_ID: item.stageId,
    CURRENCY_ID: 'BRL',
    CONTACT_ID: contatoId,
    OPENED: 'Y',
    ASSIGNED_BY_ID: RESPONSAVEL_ID,
    'UF_CRM_VALOR_BOLETO': String(valor).replace(',', '.'),   // Valor do boleto (com ponto: o PowerZap separa os #params por virgula)
    'UF_CRM_VENCIMENTO_BOLETO': item.venc,      // Vencimento do boleto
    [CAMPO_ETAPA]: tagEtapa(item.etapa),        // tag da etapa (D0, D5...)
    [CAMPO_CHAVE]: chave                         // chave da fatura (evita duplicar o card)
  } });
  if (!r.result) throw new Error('negocio: ' + JSON.stringify(r));
  return r.result;
}

// --- Anti-duplicidade: 1 card por fatura ---

// chave da fatura = telefone | valor | vencimento (ex.: 5511953449492|350.00|19/10/2026)
function chaveFatura(fone, valor, venc) {
  return fone + '|' + String(valor).replace(',', '.') + '|' + venc;
}

// ordem das etapas (da mais cedo para a mais avancada): D0, D5, D10, D15, D30
function ordemEtapas() {
  return Object.keys(REGUA).map(Number).sort(function (a, b) { return a - b; })
    .map(function (k) { return tagEtapa(REGUA[String(k)]); });
}

// procura card ABERTO desta fatura no funil. So aceita se o campo chave voltar igual (nunca confia so no filtro)
function buscarNegocioPorChave(chave) {
  const f = {}; f.CATEGORY_ID = CATEGORY_ID; f.CLOSED = 'N'; f[CAMPO_CHAVE] = chave;
  const r = chamar('crm.deal.list', { filter: f, select: ['ID', 'STAGE_ID', CAMPO_CHAVE, CAMPO_ETAPA] });
  const lista = r.result || [];
  for (let i = 0; i < lista.length; i++) if (lista[i][CAMPO_CHAVE] === chave) return lista[i];
  return null;
}

// fatura ja tem card: se avancou de etapa, MOVE o card (o robo da nova etapa dispara); senao nao faz nada
function tratarNegocioExistente(ex, item) {
  const novaTag = tagEtapa(item.etapa), atual = String(ex[CAMPO_ETAPA] || '').toUpperCase();
  if (ex.STAGE_ID === item.stageId) return 'OK ja existia negocio ' + ex.ID + ' (mesma etapa, nao reenviado)';
  const ordem = ordemEtapas();
  if (ordem.indexOf(atual) > ordem.indexOf(novaTag)) return 'OK ja existia negocio ' + ex.ID + ' (em etapa mais avancada: ' + atual + ')';
  const f = { STAGE_ID: item.stageId }; f[CAMPO_ETAPA] = novaTag;
  const r = chamar('crm.deal.update', { id: ex.ID, fields: f });
  if (!r.result) throw new Error('mover negocio ' + ex.ID + ': ' + JSON.stringify(r));
  return 'OK negocio ' + ex.ID + ' movido para ' + item.etapa;
}

function campoChaveExiste() {
  const r = chamar('crm.deal.userfield.list', { filter: { FIELD_NAME: CAMPO_CHAVE } });
  return !!(r.result && r.result.length);
}

// cria o campo "Chave da fatura" no Negocio (rodar 1 vez)
function criarCampoChave() {
  const rotulo = 'Chave da fatura';
  const r = chamar('crm.deal.userfield.add', { fields: {
    FIELD_NAME: 'CHAVE_FATURA', USER_TYPE_ID: 'string', XML_ID: 'CHAVE_FATURA',
    EDIT_FORM_LABEL: { br: rotulo, en: rotulo }, LIST_COLUMN_LABEL: { br: rotulo, en: rotulo }, LIST_FILTER_LABEL: { br: rotulo, en: rotulo }
  } });
  SpreadsheetApp.getUi().alert(r.result ? 'Campo criado: UF_CRM_CHAVE_FATURA' : 'Retorno: ' + JSON.stringify(r));
}

// Cria automaticamente, sem alerta, os campos personalizados que ainda nao existem no Negocio.
// Retorna a lista dos que nao foi possivel criar (vazia = tudo certo).
function garantirCampos() {
  const campos = [
    { nome: 'VENCIMENTO_BOLETO', rotulo: 'Vencimento do boleto' },
    { nome: 'ETAPA_REGUA', rotulo: 'Etapa da regua' },
    { nome: 'CHAVE_FATURA', rotulo: 'Chave da fatura' }
  ];
  const falhou = [];
  campos.forEach(function (c) {
    const l = chamar('crm.deal.userfield.list', { filter: { FIELD_NAME: 'UF_CRM_' + c.nome } });
    if (l.result && l.result.length) return;
    const r = chamar('crm.deal.userfield.add', { fields: {
      FIELD_NAME: c.nome, USER_TYPE_ID: 'string', XML_ID: c.nome,
      EDIT_FORM_LABEL: { br: c.rotulo, en: c.rotulo }, LIST_COLUMN_LABEL: { br: c.rotulo, en: c.rotulo }, LIST_FILTER_LABEL: { br: c.rotulo, en: c.rotulo }
    } });
    if (!r.result) falhou.push(c.rotulo);
  });
  return falhou;
}

// --- Auxiliares ---

function chamar(metodo, payload) {
  // ate 4 tentativas. So repete quando a chamada NAO chegou ao Bitrix (falha de conexao)
  // ou foi recusada por excesso de chamadas, para nao duplicar contato/card.
  for (let t = 1; t <= 4; t++) {
    try {
      const res = UrlFetchApp.fetch(WEBHOOK_URL + metodo + '.json', { method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true });
      const txt = res.getContentText();
      let j;
      try { j = JSON.parse(txt); } catch (e) { j = null; }
      if (j && j.error === 'QUERY_LIMIT_EXCEEDED' && t < 4) { Utilities.sleep(2000 * t); continue; }
      if (!j) throw new Error('resposta invalida do Bitrix (HTTP ' + res.getResponseCode() + ')');
      return j;
    } catch (e) {
      const m = String(e.message || e);
      const semConexao = /Endere[cç]o n[aã]o dispon[ií]vel|Address unavailable|DNS|timed out|Timeout|Servi[cç]o indispon[ií]vel|Service invoked too many times/i.test(m);
      if (semConexao && t < 4) { Utilities.sleep(2000 * t); continue; }
      throw e;
    }
  }
}

function buscarEtapas() {
  return (chamar('crm.dealcategory.stage.list', { id: CATEGORY_ID }).result || []).map(function (s) {
    return { id: s.STATUS_ID, nome: normalizar(s.NAME) };
  });
}

// "D-05", "D-5" e "D-5 Aviso" valem como D-5 (e D+1 nao confunde com D+14)
function normalizar(s) {
  return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '').replace(/([+-])0+(\d)/g, '$1$2');
}
function acharEtapa(etapas, chave) {
  const base = normalizar(chave);
  const re = new RegExp('^(' + base.replace(/\+/g, '\\+') + '|' + base.replace(/\+/g, '') + ')(?!\\d)');   // aceita "D+5" e "D5" (D5 nao casa com D15)
  for (let i = 0; i < etapas.length; i++) if (re.test(etapas[i].nome)) return etapas[i].id;
  return null;
}

function limparFone(v) {
  let f = String(v || '').replace(/\D/g, '');
  if (f.length === 10 || f.length === 11) f = '55' + f;
  return f;
}

function hojeUTC() {
  const p = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd').split('-');
  return Date.UTC(+p[0], +p[1] - 1, +p[2]);
}

// aceita data do Sheets ou texto dd/mm/aaaa
function lerData(v) {
  let y, m, d;
  if (v instanceof Date && !isNaN(v)) {
    const p = Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd').split('-'); y = +p[0]; m = +p[1]; d = +p[2];
  } else {
    const t = String(v || '').trim();
    const iso = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]|$)/);              // 2026-10-01 (ou 2026-10-01T00:00:00)
    const br = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);               // 01/10/2026, 01-10-2026, 01.10.2026
    if (iso) { y = +iso[1]; m = +iso[2]; d = +iso[3]; }
    else if (br) { d = +br[1]; m = +br[2]; y = +br[3]; if (y < 100) y += 2000; }
    else return null;
  }
  const utc = Date.UTC(y, m - 1, d), dt = new Date(utc);
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  const z = function (n) { return (n < 10 ? '0' : '') + n; };
  return { utc: utc, texto: z(d) + '/' + z(m) + '/' + y };
}

// "D+5" -> "D5" (tag usada no Bitrix e na planilha)
function tagEtapa(etapa) { return String(etapa || '').replace('+', '').toUpperCase(); }

// cria o campo "Etapa da regua" no Negocio (rodar 1 vez)
function criarCampoEtapa() {
  const rotulo = 'Etapa da regua';
  const r = chamar('crm.deal.userfield.add', { fields: {
    FIELD_NAME: 'ETAPA_REGUA', USER_TYPE_ID: 'string', XML_ID: 'ETAPA_REGUA',
    EDIT_FORM_LABEL: { br: rotulo, en: rotulo }, LIST_COLUMN_LABEL: { br: rotulo, en: rotulo }, LIST_FILTER_LABEL: { br: rotulo, en: rotulo }
  } });
  SpreadsheetApp.getUi().alert(r.result ? 'Campo criado: UF_CRM_ETAPA_REGUA' : 'Retorno: ' + JSON.stringify(r));
}

// cria o campo "Vencimento do boleto" no Negócio (rodar 1 vez)
function criarCampoVencimento() {
  const r = chamar('crm.deal.userfield.add', { fields: { FIELD_NAME: 'VENCIMENTO_BOLETO', USER_TYPE_ID: 'string', XML_ID: 'VENCIMENTO_BOLETO',
    EDIT_FORM_LABEL: { br: 'Vencimento do boleto', en: 'Vencimento do boleto' },
    LIST_COLUMN_LABEL: { br: 'Vencimento do boleto', en: 'Vencimento do boleto' },
    LIST_FILTER_LABEL: { br: 'Vencimento do boleto', en: 'Vencimento do boleto' } } });
  SpreadsheetApp.getUi().alert(r.result ? 'Campo criado: UF_CRM_VENCIMENTO_BOLETO' : 'Retorno: ' + JSON.stringify(r));
}


// =====================================================================
//  ORGANIZAR: reconhece colunas pelo conteudo e monta a aba "Cobranca"
// =====================================================================

function organizar() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  // De onde ler: da aba "Colar aqui" (relatorio bruto) ou, se ela estiver vazia, da propria aba "Cobranca"
  // (quando alguem colou o relatorio direto la). Nunca le a "Cobranca" se ela ja tem envios feitos (OK).
  const colar = ss.getSheetByName(ABA_COLAR);
  let bruto = colar ? colar.getDataRange().getValues() : [];
  const vazio = function (m) { return m.length < 1 || m.every(function (r) { return r.every(function (v) { return v === '' || v == null; }); }); };
  let origem = ABA_COLAR;
  if (vazio(bruto)) {
    const cob0 = ss.getSheetByName(ABA_COBRANCA);
    const b2 = cob0 ? cob0.getDataRange().getValues() : [];
    if (!vazio(b2)) {
      if (b2.some(function (r) { return /^OK/i.test(String(r[5] || '')); })) {
        ui.alert('A aba "' + ABA_COBRANCA + '" ja tem envios feitos (OK). Para um novo lote, cole o relatorio na aba "' + ABA_COLAR + '" ou limpe a "' + ABA_COBRANCA + '".');
        return;
      }
      bruto = b2; origem = ABA_COBRANCA;
    }
  }
  if (vazio(bruto)) { ui.alert('Nao achei dados. Cole o relatorio na aba "' + ABA_COLAR + '" (ou na "' + ABA_COBRANCA + '") e rode de novo.'); return; }

  // celulas com erro (ex.: telefone "+55 ..." vira #ERROR! quando colado em celula nao formatada como texto)
  let celulasErro = 0;
  bruto.forEach(function (row) { row.forEach(function (v) { if (typeof v === 'string' && /^#(ERROR|NAME\?|VALUE|REF|N\/A|DIV)/.test(v)) celulasErro++; }); });
  if (celulasErro) {
    ui.alert('Celulas com erro na aba "' + ABA_COLAR + '"',
      'Encontrei ' + celulasErro + ' celula(s) com erro (#ERROR! ou parecido). Isso acontece quando o telefone comeca com "+" e o Sheets entende como formula.\n\n' +
      'Como resolver: menu "0. Preparar aba Colar aqui", apague o conteudo e cole o relatorio de novo.', ui.ButtonSet.OK);
    return;
  }

  const det = detectarColunas(bruto);
  if (det.erro) { ui.alert('Nao consegui organizar', det.erro, ui.ButtonSet.OK); return; }

  // monta as linhas limpas
  const hoje = hojeUTC(), vistos = {}, linhas = [];
  const mapaTel = construirMapaTelefones();
  const cacheBx = {};
  const cont0 = { repetidos: 0, ignoradas: 0, suspensas: 0, viaAba: 0, viaBitrix: 0, agrupadas: 0 };
  for (let i = det.inicio; i < bruto.length; i++) {
    const r = bruto[i];
    const nome = String(r[det.nome] == null ? '' : r[det.nome]).replace(/\s+/g, ' ').trim();
    if (!nome) continue;                                                  // linha vazia ou de total/grupo

    // nao cobrar: fatura paga, cancelada ou com data de pagamento
    const statusTxt = det.status != null ? semAcento(r[det.status]) : '';
    const pagto = det.pagto != null ? r[det.pagto] : '';
    if (/\bPAGA\b|CANCEL|QUITAD|BAIXAD/.test(statusTxt) || (pagto !== '' && pagto != null)) { cont0.ignoradas++; continue; }
    // nao cobrar: UC cancelada/suspensa
    if (r.some(function (v) { return /SUSPENS/.test(semAcento(v)); })) { cont0.suspensas++; continue; }

    const valor = normalizarValor(r[det.valor]);
    const ven = lerData(r[det.venc]);
    let fone = det.fone != null ? limparFone(r[det.fone]) : '';
    let erro = '';
    if (!valor) erro = 'valor invalido: ' + r[det.valor];
    else if (!ven) erro = 'data de vencimento invalida: ' + r[det.venc];
    if (!erro && !/^55\d{10,11}$/.test(fone)) {
      // telefone nao veio no relatorio: procura na aba "Telefones" e depois no Bitrix (pelo nome)
      const uc = det.uc != null ? r[det.uc] : '';
      let ach = achaNaAba(mapaTel, nome, uc);
      if (ach) cont0.viaAba++;
      else {
        if (!(nome in cacheBx)) cacheBx[nome] = buscarTelefoneBitrix(nome);
        ach = cacheBx[nome];
        if (ach) cont0.viaBitrix++;
      }
      fone = ach || '';
      if (!fone) erro = 'telefone nao encontrado';
    }
    if (!erro) {
      const k = fone + '|' + valor + '|' + ven.texto;
      if (vistos[k]) { cont0.repetidos++; continue; }
      vistos[k] = true;
    }
    const etp = (!erro && ven) ? etapaPorDias(Math.round((hoje - ven.utc) / 86400000)) : '';   // etapa de hoje (ou vazio se fora da regua)
    linhas.push([nome, fone ? String(fone) : '', valor, ven ? ven.texto : String(r[det.venc] || ''), etp ? tagEtapa(etp) : '', erro ? 'ERRO: ' + erro : '']);
  }
  const repetidos = cont0.repetidos;
  // Cliente com varias faturas na regua no mesmo dia recebe UMA mensagem so: a da etapa mais avancada
  // (a mensagem fala apenas "fatura em aberto"). As demais ficam como AGRUPADO e nao sao enviadas.
  // Para enviar uma mensagem por fatura, troque AGRUPAR_POR_CLIENTE para false.
  if (AGRUPAR_POR_CLIENTE) {
    const melhor = {};
    linhas.forEach(function (l, idx) {
      if (!l[4] || l[5]) return;
      const k = semAcento(l[0]);
      const dias = Math.round((hoje - lerData(l[3]).utc) / 86400000);
      if (!(k in melhor) || dias > melhor[k].dias) melhor[k] = { idx: idx, dias: dias };
    });
    linhas.forEach(function (l, idx) {
      if (!l[4] || l[5]) return;
      const m = melhor[semAcento(l[0])];
      if (m.idx !== idx) { cont0.agrupadas++; l[5] = 'AGRUPADO: enviado so ' + linhas[m.idx][4] + ' hoje'; l[4] = ''; }
    });
  }
  if (!linhas.length) { ui.alert('Nenhuma linha de dados encontrada.'); return; }

  // mostra o que foi entendido e pede confirmacao
  const ex = linhas.filter(function (l) { return !l[5]; })[0] || linhas[0];
  const nomeCol = function (c) { return 'coluna ' + colLetra(c) + (det.cab ? ' ("' + det.cab[c] + '")' : ''); };
  let msg = 'Entendi assim:\n' +
    '• Nome: ' + nomeCol(det.nome) + '\n• Telefone: ' + (det.fone != null ? nomeCol(det.fone) : 'nao veio no relatorio (busco na aba "Telefones" e no Bitrix, pelo nome)') + '\n• Valor: ' + nomeCol(det.valor) + '\n• Vencimento: ' + nomeCol(det.venc) +
    '\n\nExemplo da primeira linha:\n' + ex[0] + ' | ' + ex[1] + ' | R$ ' + ex[2] + ' | ' + ex[3];
  if (det.avisos.length) msg += '\n\nAtencao:\n- ' + det.avisos.join('\n- ');
  msg += '\n\nEsta correto? (Se nao, clique em Nao e confira a aba "' + ABA_COLAR + '")';
  if (ui.alert('Conferir colunas', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // grava a aba Cobranca
  let cob = ss.getSheetByName(ABA_COBRANCA);
  if (!cob) cob = ss.insertSheet(ABA_COBRANCA);
  cob.clear();
  cob.getRange(1, 1, 1, 6).setValues([['NOME', 'TELEFONE', 'VALOR', 'VENCIMENTO', 'ETAPA', 'STATUS']]).setFontWeight('bold');
  cob.getRange(2, 2, linhas.length, 3).setNumberFormat('@');          // texto, para nao virar numero/data
  cob.getRange(2, 1, linhas.length, 6).setValues(linhas);
  cob.setFrozenRows(1);
  cob.autoResizeColumns(1, 6);

  // resumo por etapa
  const cont = {}; let fora = 0, erros = 0;
  linhas.forEach(function (l) {
    if (/^AGRUPADO/.test(l[5])) return;
    if (l[5]) { erros++; return; }
    const ven = lerData(l[3]);
    const dias = Math.round((hoje - ven.utc) / 86400000);
    const et = etapaPorDias(dias);
    if (et) cont[et] = (cont[et] || 0) + 1; else fora++;
  });
  const resumo = Object.keys(REGUA).map(function (k) { return REGUA[k]; })
    .filter(function (e) { return cont[e]; }).map(function (e) { return e + ': ' + cont[e]; }).join(' | ');
  ss.setActiveSheet(cob);
  ui.alert('Organizado!\n\n' + linhas.length + ' linha(s) na aba "' + ABA_COBRANCA + '"' +
    (repetidos ? ' (' + repetidos + ' repetida(s) removida(s))' : '') + '\n' +
    (resumo ? 'Para enviar hoje: ' + resumo : 'Nenhuma linha cai na regua hoje.') +
    (fora ? '\nFora da regua hoje (nao serao enviadas): ' + fora : '') +
    (cont0.agrupadas ? '\nFaturas agrupadas (cliente recebe so a mensagem da etapa mais avancada): ' + cont0.agrupadas + ' linha(s)' : '') +
    (cont0.ignoradas ? '\nIgnoradas (paga/cancelada): ' + cont0.ignoradas : '') +
    (cont0.suspensas ? '\nIgnoradas (UC cancelada/suspensa): ' + cont0.suspensas : '') +
    ((cont0.viaAba || cont0.viaBitrix) ? '\nTelefone achado fora do relatorio: ' + (cont0.viaAba + cont0.viaBitrix) + ' (aba Telefones: ' + cont0.viaAba + ', Bitrix: ' + cont0.viaBitrix + ')' : '') +
    (erros ? '\nCom erro (corrija ou apague): ' + erros : '') +
    '\n\nProximo passo: menu "2. Enviar Registros".');
}

function colLetra(c) { let s = '', n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// ---------- reconhecimento de colunas ----------
function pareceFone(v) {
  if (v === '' || v == null || v instanceof Date) return false;
  const t = String(v).trim();
  if (/^\d{3}\.\d{3}\.\d{3}-\d{2}$/.test(t) || /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/.test(t)) return false;   // CPF/CNPJ formatado
  if (/[A-Za-z@]/.test(t)) return false;
  const d = t.replace(/\D/g, '');
  if (d.length === 14) return false;                                                                           // CNPJ sem formatacao
  return /^(55)?[1-9]\d9\d{8}$/.test(d) || /^(55)?[1-9]\d[2-5]\d{7}$/.test(d);
}
function pareceData(v) {
  if (v instanceof Date && !isNaN(v)) return true;
  const t = String(v == null ? '' : v).trim();
  return /^(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})/.test(t) && !!lerData(t);
}
function pareceValor(v) {
  if (v === '' || v == null || v instanceof Date) return false;
  if (typeof v === 'number') return v > 0 && v < 10000000 && !(v >= 1000000000);
  const t = String(v).trim();
  return /^(R\$\s*)?\d{1,3}(\.\d{3})*,\d{2}$/.test(t) || /^(R\$\s*)?\d+,\d{1,2}$/.test(t) || /^(R\$\s*)?\d+\.\d{1,2}$/.test(t);
}
function pareceNome(v) {
  if (v === '' || v == null || v instanceof Date || typeof v === 'number') return false;
  const t = String(v).trim();
  if (t.length < 3 || /@/.test(t)) return false;
  const letras = (t.match(/[A-Za-zÀ-ÿ]/g) || []).length;
  return letras / t.length >= 0.6;
}

const DICAS = {
  nome: /NOME|CLIENTE|RAZAO|PAGADOR|SACADO|TITULAR|ASSOCIADO/,
  fone: /CELULAR|WHATS|FONE|TELEFONE|CONTATO|MOVEL/,
  valor: /VALOR|TOTAL|SALDO|DEBITO|PRINCIPAL/,
  venc: /VENC/
};
const EVITAR_DATA = /EMISS|CRIA|PAGAMENTO|PAGO|COMPET|BAIXA|REFER/;

function detectarColunas(m) {
  const norm = function (x) { return String(x == null ? '' : x).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); };
  const ncol = Math.max.apply(null, m.map(function (r) { return r.length; }));
  // acha a 1a linha com dados (telefone/data/valor); o cabecalho e a linha logo acima dela (se tiver 3+ celulas preenchidas)
  const temDado = function (r) { return r.some(function (v) { return pareceFone(v) || pareceData(v) || pareceValor(v); }); };
  let inicio = 0;
  while (inicio < m.length && inicio < 30 && !temDado(m[inicio])) inicio++;
  if (inicio >= m.length || inicio >= 30) return { erro: 'Nao achei linhas com telefone, data ou valor. Confira a aba "' + ABA_COLAR + '".' };
  let cab = null;
  if (inicio > 0) {
    const h = m[inicio - 1];
    if (h.filter(function (x) { return String(x == null ? '' : x).trim() !== ''; }).length >= 3)
      cab = h.map(function (x) { return String(x == null ? '' : x).replace(/\s+/g, ' ').trim(); });
  }
  const dados = m.slice(inicio, inicio + 200);
  if (!dados.length) return { erro: 'Achei so o cabecalho, sem linhas de dados.' };

  const pontos = function (c, fn) {
    let n = 0, ok = 0;
    dados.forEach(function (r) { const v = r[c]; if (v === '' || v == null) return; n++; if (fn(v)) ok++; });
    return n ? ok / n : 0;
  };
  const info = [];
  for (let c = 0; c < ncol; c++) {
    info.push({ c: c, h: cab ? norm(cab[c]) : '', data: pontos(c, pareceData), fone: pontos(c, pareceFone), valor: pontos(c, pareceValor), nome: pontos(c, pareceNome) });
  }
  const usadas = {}, avisos = [], res = { cab: cab, inicio: inicio, avisos: avisos };

  function escolher(chave, cands, dica, desempate, rotulo) {
    cands = cands.filter(function (x) { return !usadas[x.c]; });
    if (!cands.length) return null;
    const comDica = cands.filter(function (x) { return dica && dica.test(x.h); });
    let pick;
    if (comDica.length) pick = comDica[0];
    else {
      cands.sort(desempate);
      pick = cands[0];
      if (cands.length > 1) avisos.push('Havia mais de uma coluna possivel para ' + rotulo + '; usei a coluna ' + colLetra(pick.c) + '. Confira o exemplo abaixo.');
    }
    usadas[pick.c] = true;
    return pick.c;
  }

  const datas = info.filter(function (x) { return x.data >= 0.6; });
  const datasBoas = datas.filter(function (x) { return !EVITAR_DATA.test(x.h); });
  res.venc = escolher('venc', datasBoas.length ? datasBoas : datas, DICAS.venc, function (a, b) { return b.c - a.c; }, 'vencimento');
  res.fone = escolher('fone', info.filter(function (x) { return x.fone >= 0.6; }), DICAS.fone, function (a, b) { return b.fone - a.fone || a.c - b.c; }, 'telefone');
  res.valor = escolher('valor', info.filter(function (x) { return x.valor >= 0.6 && x.data < 0.6 && x.fone < 0.6; }), DICAS.valor, function (a, b) { return b.valor - a.valor || a.c - b.c; }, 'valor');
  res.nome = escolher('nome', info.filter(function (x) { return x.nome >= 0.6; }), DICAS.nome, function (a, b) { return b.nome - a.nome || a.c - b.c; }, 'nome');

  // colunas extras, so pelo cabecalho
  const achaCab = function (re) { for (let c = 0; c < info.length; c++) if (info[c].h && re.test(info[c].h)) return c; return null; };
  res.status = achaCab(/STATUS|SITUACAO/);
  res.pagto = achaCab(/PAGAMENTO|PAGO EM|BAIXA/);
  res.uc = achaCab(/UNIDADE|CODIGO|\bUC\b/);
  const falt = [];
  if (res.nome == null) falt.push('nome');
  if (res.valor == null) falt.push('valor'); if (res.venc == null) falt.push('data de vencimento');
  if (falt.length) return { erro: 'Nao consegui reconhecer a coluna de: ' + falt.join(', ') + '.\nConfira se a aba "' + ABA_COLAR + '" tem esses dados em colunas separadas e sem celulas mescladas.' };
  return res;
}

function normalizarValor(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v > 0 ? v.toFixed(2).replace('.', ',') : '';
  let t = String(v).replace(/R\$\s*/gi, '').trim();
  if (/^\d{1,3}(\.\d{3})+(,\d{1,2})?$/.test(t)) t = t.replace(/\./g, '');      // 1.761,31 -> 1761,31
  else if (/^\d+\.\d{1,2}$/.test(t)) t = t.replace('.', ',');                  // 1761.31 -> 1761,31
  if (/^\d+$/.test(t)) t += ',00';                                             // 150 -> 150,00
  if (/^\d+,\d$/.test(t)) t += '0';                                            // 89,9 -> 89,90
  return /^\d+,\d{2}$/.test(t) ? t : '';
}

// ---------- telefone fora do relatorio ----------
function semAcento(v) {
  return String(v == null ? '' : v).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}

// aba opcional "Telefones": qualquer tabela com um telefone por linha e o nome e/ou codigo da UC/CPF nas outras colunas
function construirMapaTelefones() {
  const mapa = {};
  const sh = SpreadsheetApp.getActive().getSheetByName('Telefones');
  if (!sh || sh.getLastRow() < 1) return mapa;
  sh.getDataRange().getValues().forEach(function (row) {
    let fone = '';
    row.forEach(function (v) { if (!fone && pareceFone(v)) fone = limparFone(v); });
    if (!fone) return;
    row.forEach(function (v) {
      if (v === '' || v == null || pareceFone(v)) return;
      const t = semAcento(v);
      if (t.length >= 4) mapa['N:' + t.replace(/[^A-Z0-9]+/g, ' ').trim()] = fone;
      const d = String(v).replace(/\D/g, '');
      if (d.length >= 8) mapa['D:' + d] = fone;
    });
  });
  return mapa;
}
function achaNaAba(mapa, nome, uc) {
  const d = String(uc == null ? '' : uc).replace(/\D/g, '');
  if (d.length >= 8 && mapa['D:' + d]) return mapa['D:' + d];
  return mapa['N:' + semAcento(nome).replace(/[^A-Z0-9]+/g, ' ').trim()] || null;
}

// procura o contato no Bitrix pelo nome exato e devolve o 1o celular valido
function buscarTelefoneBitrix(nome) {
  try {
    const r = chamar('crm.contact.list', { filter: { '=NAME': nome }, select: ['ID', 'NAME', 'PHONE'] });
    const lista = r.result || [];
    for (let i = 0; i < lista.length; i++) {
      const fones = lista[i].PHONE || [];
      for (let k = 0; k < fones.length; k++) {
        const f = limparFone(fones[k].VALUE);
        if (/^55\d{10,11}$/.test(f)) return f;
      }
    }
  } catch (e) { /* sem acesso: segue sem telefone */ }
  return null;
}


// formata a aba "Colar aqui" inteira como TEXTO SIMPLES, para o Sheets nao transformar "+55 11 9..." em formula
function prepararAbaColar() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ABA_COLAR);
  if (!sh) sh = ss.insertSheet(ABA_COLAR);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
  ss.setActiveSheet(sh);
  ui.alert('Aba "' + ABA_COLAR + '" pronta.\n\nApague o que tiver nela e cole o relatorio (de preferencia com Ctrl+Shift+V, "colar somente valores"). ' +
    'Depois use o menu "1. Organizar o que colei".');
}


// =====================================================================
//  TESTAR CONFIGURACAO: confere webhook, funil, etapas e usuario antes do 1o envio
// =====================================================================
function testarConfiguracao() {
  const ui = SpreadsheetApp.getUi();
  const falta = [];
  if (/COLE_AQUI/.test(WEBHOOK_URL)) falta.push('WEBHOOK_URL');
  if (/COLE_AQUI/.test(CATEGORY_ID)) falta.push('CATEGORY_ID');
  if (/COLE_AQUI/.test(RESPONSAVEL_ID)) falta.push('RESPONSAVEL_ID');
  if (falta.length) { ui.alert('Falta preencher no topo do script: ' + falta.join(', ') + '.'); return; }
  if (!/\/$/.test(WEBHOOK_URL)) { ui.alert('O WEBHOOK_URL precisa terminar com "/".'); return; }

  const linhas = [];
  try {
    const cats = chamar('crm.dealcategory.list', {});
    if (cats.error) { ui.alert('Webhook recusado pelo Bitrix: ' + (cats.error_description || cats.error) + '\n\nConfira o endereco e a permissao de CRM.'); return; }
    const cat = (cats.result || []).filter(function (c) { return String(c.ID) === String(CATEGORY_ID); })[0];
    linhas.push(cat ? 'Funil ' + CATEGORY_ID + ': OK ("' + cat.NAME + '")' : 'Funil ' + CATEGORY_ID + ': NAO ENCONTRADO. Confira o ID no endereco do funil.');

    const etapas = buscarEtapas();
    Object.keys(REGUA).forEach(function (k) {
      const achou = !!acharEtapa(etapas, REGUA[k]);
      linhas.push('Etapa ' + REGUA[k] + ': ' + (achou ? 'OK' : 'NAO ENCONTRADA no funil'));
    });

    const u = chamar('user.get', { ID: RESPONSAVEL_ID });
    if (u.result && u.result.length) linhas.push('Responsavel ' + RESPONSAVEL_ID + ': OK (' + u.result[0].NAME + ' ' + u.result[0].LAST_NAME + ')');
    else linhas.push('Responsavel ' + RESPONSAVEL_ID + ': nao consegui confirmar (o webhook pode nao ter permissao de usuarios). Confira o ID no perfil.');

    const faltam = ['VENCIMENTO_BOLETO', 'ETAPA_REGUA', 'CHAVE_FATURA'].filter(function (n) {
      const l = chamar('crm.deal.userfield.list', { filter: { FIELD_NAME: 'UF_CRM_' + n } });
      return !(l.result && l.result.length);
    });
    linhas.push(faltam.length ? 'Campos a criar no 1o envio: ' + faltam.join(', ') : 'Campos personalizados: OK');
  } catch (e) {
    ui.alert('Nao consegui falar com o Bitrix: ' + e.message + '\n\nConfira o WEBHOOK_URL e a conexao.'); return;
  }
  ui.alert('Teste de configuracao', linhas.join('\n'), ui.ButtonSet.OK);
}
