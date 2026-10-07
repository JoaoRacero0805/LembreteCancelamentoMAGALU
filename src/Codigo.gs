/**
 * Lembrete diário de cancelamentos pendentes
 * Apenas LÊ a planilha e envia por e-mail os pedidos que não estão como "Cancelado".
 * Envio automático às 16:30 em dias úteis (rode criarGatilhos UMA vez).
 *
 * Dados sensíveis NÃO ficam no código. Configure em
 * Configurações do projeto > Propriedades do script:
 *   EMAILS_DESTINO    (obrigatória) e-mails dos supervisores, separados por vírgula
 *   NOME_TIME         (opcional)    nome do time exibido no cabeçalho do e-mail
 *   ROTULO_PROTOCOLO  (opcional)    título da coluna de protocolo (padrão: "Protocolo")
 */

const CONFIG = {
  NOME_ABA: 'Página1',
  LINHA_INICIAL: 3,                      // dados começam na linha 3
  HORA_ENVIO: 16,
  MINUTO_ENVIO: 30,                      // envio às 16:30
  ENVIAR_FIM_DE_SEMANA: false,
  ENVIAR_SE_VAZIO: true,                 // manda "nenhum pendente" quando não houver nada
  COL: { PEDIDO: 1, PROTOCOLO: 2, MOTIVO: 3, DATA: 4, STATUS: 5 },

  // Logo no topo do e-mail (versão branca, para fundo azul). Se a imagem não carregar,
  // troque pela URL de outro logo ou deixe '' para usar o nome "magalu" escrito em texto.
  LOGO_URL: 'https://companieslogo.com/img/orig/MGLU3.SA_BIG.D-e463d6bc.png',

  // Identidade visual Magalu
  COR_AZUL: '#0086FF',                   // azul Magalu
  COR_AZUL_ESCURO: '#0A4FA3',
  COR_TEXTO: '#1F2937',
  COR_TEXTO_SUAVE: '#6B7280',
  COR_FUNDO: '#F2F5F9',
  FAIXA_ARCO_IRIS: ['#F5B400', '#F57C00', '#E53935', '#E4007C', '#8E24AA', '#0086FF', '#00A650'],

  // Prioridade
  COR_HOJE: '#FFF6D6',                   // amarelo: solicitado hoje, ainda no prazo
  COR_HOJE_FORTE: '#F5B400',
  COR_ATRASADO: '#FDE4E4',               // vermelho: solicitado antes de hoje e não cancelado
  COR_ATRASADO_FORTE: '#E53935',
};

/* ===================== AGENDAMENTO ===================== */

// Rode UMA vez. Cria um gatilho diário (entre 15h e 16h) que agenda o envio exato das 16:30.
function criarGatilhos() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['enviarLembrete', 'agendarEnvioDoDia'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('agendarEnvioDoDia')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.HORA_ENVIO - 1)
    .inTimezone('America/Sao_Paulo')
    .create();
}

// Executada automaticamente todo dia: agenda um disparo único para hoje às 16:30.
function agendarEnvioDoDia() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarLembrete')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const alvo = new Date();
  alvo.setHours(CONFIG.HORA_ENVIO, CONFIG.MINUTO_ENVIO, 0, 0);
  if (alvo <= new Date()) return;        // segurança: se já passou do horário, não agenda

  ScriptApp.newTrigger('enviarLembrete').timeBased().at(alvo).create();
}

/* ======================= E-MAIL ======================= */

// Chamada pelo agendamento (NÃO execute manualmente: envia para a supervisão)
function enviarLembrete() { enviar_(false); }

// Use esta para testar: envia só para você, em qualquer dia
function testarEmail() { enviar_(true); }

function enviar_(teste) {
  const agora = new Date();
  const dia = agora.getDay();
  if (!teste && !CONFIG.ENVIAR_FIM_DE_SEMANA && (dia === 0 || dia === 6)) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.NOME_ABA);
  if (!aba) throw new Error('Aba "' + CONFIG.NOME_ABA + '" não encontrada.');

  const pendentes = lerPendentes_(aba, agora);
  if (!pendentes.length && !CONFIG.ENVIAR_SE_VAZIO && !teste) return;

  const fuso = ss.getSpreadsheetTimeZone();
  const dataHoje = Utilities.formatDate(agora, fuso, 'dd/MM/yyyy');
  const assunto = (teste ? '[TESTE] ' : '') + (pendentes.length
    ? `[Cancelamentos] ${pendentes.length} pedido(s) pendente(s) – ${dataHoje}`
    : `[Cancelamentos] Nenhum pendente – ${dataHoje}`);

  MailApp.sendEmail({
    to: teste ? Session.getActiveUser().getEmail() : obterDestinatarios_(),
    subject: assunto,
    htmlBody: montarEmail_(pendentes, ss, aba, fuso, dataHoje),
    name: 'Lembrete de Cancelamentos',
  });
}

function lerPendentes_(aba, agora) {
  const C = CONFIG.COL;
  const ultima = aba.getLastRow();
  if (ultima < CONFIG.LINHA_INICIAL) return [];

  const dados = aba.getRange(CONFIG.LINHA_INICIAL, 1, ultima - CONFIG.LINHA_INICIAL + 1, C.STATUS).getValues();
  const hoje = zerarHora_(agora);

  return dados
    .map((l, i) => ({
      linha: CONFIG.LINHA_INICIAL + i,
      pedido: String(l[C.PEDIDO - 1]).trim(),
      protocolo: String(l[C.PROTOCOLO - 1]).trim(),
      motivo: String(l[C.MOTIVO - 1]).trim(),
      data: l[C.DATA - 1],
      status: normalizar_(l[C.STATUS - 1]),
    }))
    .filter(p => p.pedido && p.status !== 'cancelado')   // tudo que não for "Cancelado" é pendente
    .map(p => {
      p.dias = p.data instanceof Date ? Math.round((hoje - zerarHora_(p.data)) / 86400000) : null;
      return p;
    })
    // mais antigos primeiro; sem data válida vão para o final
    .sort((a, b) => (b.dias == null ? -Infinity : b.dias) - (a.dias == null ? -Infinity : a.dias));
}

/* =================== VISUAL DO E-MAIL (MAGALU) =================== */

function montarEmail_(pendentes, ss, aba, fuso, dataHoje) {
  const K = CONFIG;
  const link = ss.getUrl() + '#gid=' + aba.getSheetId();
  const atrasados = pendentes.filter(p => p.dias !== null && p.dias >= 1).length;
  const deHoje = pendentes.filter(p => p.dias !== null && p.dias <= 0).length;
  const semData = pendentes.length - atrasados - deHoje;
  const nomeTime = propriedade_('NOME_TIME', '');

  const resumo = pendentes.length
    ? `${pendentes.length} pedido(s) aguardando cancelamento, ${atrasados} atrasado(s).`
    : 'Nenhum pedido pendente de cancelamento hoje.';

  let conteudo;
  if (!pendentes.length) {
    conteudo = `
      <tr><td style="padding:40px 32px;text-align:center">
        <div style="font-size:44px;line-height:1">🎉</div>
        <div style="font-size:20px;font-weight:bold;color:${K.COR_TEXTO};padding-top:12px">Tudo em dia!</div>
        <div style="font-size:14px;color:${K.COR_TEXTO_SUAVE};padding-top:6px">
          Nenhum pedido pendente de cancelamento hoje (${dataHoje}).</div>
      </td></tr>`;
  } else {
    conteudo = `
      <tr><td style="padding:28px 32px 8px">
        <div style="font-size:20px;font-weight:bold;color:${K.COR_TEXTO}">Olá, supervisão! 👋</div>
        <div style="font-size:14px;color:${K.COR_TEXTO_SUAVE};padding-top:6px;line-height:20px">
          Ainda há <b style="color:${K.COR_TEXTO}">${pendentes.length}</b> pedido(s) aguardando cancelamento hoje (${dataHoje}).
          Clique no número do pedido para abrir a linha na planilha.</div>
      </td></tr>
      <tr><td style="padding:16px 24px 8px">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
          ${cartao_('Pendentes', pendentes.length, K.COR_AZUL, '#E6F3FF')}
          ${cartao_('Atrasados', atrasados, K.COR_ATRASADO_FORTE, K.COR_ATRASADO)}
          ${cartao_('De hoje', deHoje, '#B7791F', K.COR_HOJE)}
        </tr></table>
      </td></tr>
      <tr><td style="padding:16px 32px 0">${tabela_(pendentes, link, fuso)}</td></tr>
      <tr><td style="padding:14px 32px 0;font-size:12px;color:${K.COR_TEXTO_SUAVE};line-height:18px">
        ${marcador_(K.COR_ATRASADO_FORTE)} <b>Atrasado</b>: solicitado antes de hoje e ainda não cancelado<br>
        ${marcador_(K.COR_HOJE_FORTE)} <b>Hoje</b>: cancelar até o fim do expediente
        ${semData ? `<br>${marcador_('#9CA3AF')} <b>Sem data</b>: data da solicitação inválida ou vazia na planilha` : ''}
      </td></tr>`;
  }

  return `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:${K.COR_FUNDO}">
  <div style="display:none;max-height:0;overflow:hidden;opacity:0">${esc_(resumo)}</div>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:${K.COR_FUNDO}">
    <tr><td align="center" style="padding:24px 12px">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
             style="max-width:680px;background:#ffffff;border-radius:12px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;border:1px solid #E5E7EB">
        <tr><td style="background:${K.COR_AZUL};padding:22px 32px">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
            <td style="vertical-align:middle">${logo_()}</td>
            <td align="right" style="vertical-align:middle;color:#ffffff;font-size:12px;line-height:17px">
              <b style="font-size:14px">Cancelamentos pendentes</b><br>${nomeTime ? esc_(nomeTime) + ' · ' : ''}${dataHoje}</td>
          </tr></table>
        </td></tr>
        <tr><td style="padding:0;line-height:0;font-size:0">${faixaArcoIris_()}</td></tr>
        ${conteudo}
        <tr><td align="center" style="padding:26px 32px 30px">${botao_(link, 'Abrir a planilha')}</td></tr>
        <tr><td style="padding:0;line-height:0;font-size:0">${faixaArcoIris_()}</td></tr>
        <tr><td style="background:#F9FAFB;padding:16px 32px;font-size:11px;color:#9CA3AF;line-height:16px;text-align:center">
          E-mail automático enviado pela planilha
          <a href="${link}" style="color:${K.COR_AZUL};text-decoration:none">${esc_(ss.getName())}</a>.<br>
          Para sair da lista, marque o pedido como <b>Cancelado</b> na planilha.
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}

function tabela_(pendentes, link, fuso) {
  const K = CONFIG;
  const TH = `padding:10px 12px;text-align:left;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:#ffffff;background:${K.COR_AZUL_ESCURO}`;
  const TD = 'padding:11px 12px;font-size:13px;color:' + K.COR_TEXTO + ';border-bottom:1px solid #EEF0F3';

  const linhas = pendentes.map(p => {
    let fundo = '#ffffff', faixa = '#D1D5DB', selo = selo_('sem data', '#6B7280', '#F3F4F6');
    if (p.dias !== null && p.dias >= 1) {
      fundo = K.COR_ATRASADO; faixa = K.COR_ATRASADO_FORTE;
      selo = selo_(p.dias + (p.dias === 1 ? ' dia' : ' dias'), '#ffffff', K.COR_ATRASADO_FORTE);
    } else if (p.dias !== null) {
      fundo = K.COR_HOJE; faixa = K.COR_HOJE_FORTE;
      selo = selo_('hoje', '#5C4400', K.COR_HOJE_FORTE);
    }
    const data = p.data instanceof Date ? Utilities.formatDate(p.data, fuso, 'dd/MM/yyyy') : esc_(p.data || '—');
    return `<tr style="background:${fundo}">
      <td style="${TD};border-left:4px solid ${faixa}">
        <a href="${link}&range=A${p.linha}" style="color:${K.COR_AZUL};font-weight:bold;text-decoration:none">${esc_(p.pedido)}</a></td>
      <td style="${TD}">${esc_(p.protocolo || '—')}</td>
      <td style="${TD}">${esc_(p.motivo || '—')}</td>
      <td style="${TD};white-space:nowrap">${data}</td>
      <td style="${TD};text-align:center">${selo}</td>
    </tr>`;
  }).join('');

  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
      style="border-collapse:separate;border-spacing:0;border:1px solid #E5E7EB;border-radius:8px;overflow:hidden">
    <tr>
      <th style="${TH}">Pedido</th><th style="${TH}">${esc_(propriedade_('ROTULO_PROTOCOLO', 'Protocolo'))}</th><th style="${TH}">Motivo</th>
      <th style="${TH}">Solicitado em</th><th style="${TH};text-align:center">Prazo</th>
    </tr>
    ${linhas}
  </table>`;
}

function logo_() {
  const K = CONFIG;
  if (K.LOGO_URL) {
    return `<img src="${K.LOGO_URL}" alt="Magalu" height="34"
      style="display:block;height:34px;width:auto;border:0;outline:none;color:#ffffff;font-size:24px;font-weight:bold">`;
  }
  return `<span style="color:#ffffff;font-size:28px;font-weight:bold;letter-spacing:-1px">magalu</span>`;
}

function faixaArcoIris_() {
  const cores = CONFIG.FAIXA_ARCO_IRIS;
  const largura = (100 / cores.length).toFixed(2);
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>${
    cores.map(c => `<td width="${largura}%" height="4" style="background:${c};height:4px;line-height:4px;font-size:0">&nbsp;</td>`).join('')
  }</tr></table>`;
}

function cartao_(titulo, valor, corForte, corFundo) {
  return `<td width="33%" style="padding:0 8px">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
           style="background:${corFundo};border-radius:10px;border-top:4px solid ${corForte}">
      <tr><td style="padding:14px 16px">
        <div style="font-size:28px;font-weight:bold;color:${corForte};line-height:30px">${valor}</div>
        <div style="font-size:12px;color:${CONFIG.COR_TEXTO_SUAVE};padding-top:2px">${titulo}</div>
      </td></tr>
    </table>
  </td>`;
}

function selo_(texto, cor, fundo) {
  return `<span style="display:inline-block;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:bold;white-space:nowrap;color:${cor};background:${fundo}">${texto}</span>`;
}

function marcador_(cor) {
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${cor};vertical-align:middle"></span>`;
}

function botao_(url, texto) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
    <td style="background:${CONFIG.COR_AZUL};border-radius:999px">
      <a href="${url}" style="display:inline-block;padding:13px 30px;color:#ffffff;font-size:14px;font-weight:bold;text-decoration:none">${texto}</a>
    </td></tr></table>`;
}

/* ======================= UTILITÁRIOS ======================= */

function obterDestinatarios_() {
  const valor = propriedade_('EMAILS_DESTINO', '');
  if (!valor) {
    throw new Error('Configure a propriedade de script EMAILS_DESTINO com os e-mails de destino.');
  }
  return valor.split(',').map(e => e.trim()).filter(Boolean).join(',');
}

function propriedade_(nome, padrao) {
  const valor = PropertiesService.getScriptProperties().getProperty(nome);
  return valor && valor.trim() ? valor.trim() : padrao;
}

function normalizar_(s) {
  return String(s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function zerarHora_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
