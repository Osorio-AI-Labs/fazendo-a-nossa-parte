// ==========================================
// CONFIGURAÇÕES GERAIS
// ==========================================
// Insira aqui o e-mail exato que aparece no seu dropdown do Gmail
const EMAIL_REMETENTE = "social@osorioailabs.com.br";

// 1. Função que recebe os dados do formulário do site (COM LIMITE DE 100 E-MAILS)
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataHora = new Date();
  var hojeStr = dataHora.toLocaleDateString('pt-BR');

  // ==========================================
  // LÓGICA DE LIMITE DIÁRIO DE E-MAILS (100)
  // ==========================================
  var limiteDiario = 100;
  var enviosHoje = 0;

  // Lê a folha de cálculo para contar quantos envios ocorreram hoje
  if (sheet.getLastRow() > 1) {
    var dadosData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < dadosData.length; i++) {
      if (dadosData[i][0] instanceof Date) {
        var dataLinha = dadosData[i][0].toLocaleDateString('pt-BR');
        if (dataLinha === hojeStr) {
          enviosHoje++;
        }
      }
    }
  }
  // ==========================================

  // Captura os dados do formulário
  var nome = e.parameter.nome;
  var email = e.parameter.email;
  var whatsapp = e.parameter.whatsapp;
  var cidade = e.parameter.cidade;
  var linkedin = e.parameter.linkedin;
  var foto = e.parameter.foto;
  var cargo = e.parameter.cargo;
  var experiencia = e.parameter.experiencia;
  var tags = e.parameter.tags;
  var resumo = e.parameter.resumo;
  var historico = e.parameter.historico;
  var formacao = e.parameter.formacao;
  var competencias = e.parameter.competencias;

  // Verifica se o limite foi atingido
  if (enviosHoje < limiteDiario) {
    // ABAIXO DO LIMITE: Grava e envia o e-mail automático
    var statusInicial = "Pendente";
    sheet.appendRow([dataHora, statusInicial, nome, email, whatsapp, cidade, linkedin, foto, cargo, experiencia, tags, resumo, historico, formacao, competencias]);

    // Dispara o e-mail
    enviarEmailConfirmacao(nome, email);

  } else {
    // ACIMA DO LIMITE: Grava para não perder o talento, mas NÃO envia o e-mail
    var statusLimite = "Pendente (Limite Diário)";
    sheet.appendRow([dataHora, statusLimite, nome, email, whatsapp, cidade, linkedin, foto, cargo, experiencia, tags, resumo, historico, formacao, competencias]);
  }

  return ContentService.createTextOutput("Sucesso!").setMimeType(ContentService.MimeType.TEXT);
}

// 2. Função que monitoriza alterações (Aprovação/Recusa)
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  if (range.getColumn() == 2 && range.getRow() > 1) {
    var novoStatus = e.value;
    var linha = range.getRow();
    var nome = sheet.getRange(linha, 3).getValue();
    var email = sheet.getRange(linha, 4).getValue();

    if (novoStatus === "Aprovado") {
      enviarEmailAprovacao(nome, email);
    } else if (novoStatus === "Recusado") {
      enviarEmailRecusa(nome, email);
    }
  }
}

// 3. E-mail de Recebimento (Pendente)
function enviarEmailConfirmacao(nome, email) {
  var assunto = "Confirmação de Receção: Currículo Registado";
  var corpoHtml = gerarTemplateEmail(
    "Receção Registada",
    nome,
    "Os nossos sistemas de inteligência registaram com sucesso o envio do seu currículo para a iniciativa <strong>Fazendo a Nossa Parte</strong>.",
    "Os seus dados foram recebidos de forma segura e encaminhados imediatamente para a nossa equipa de especialistas. Para garantir a melhor apresentação do seu perfil ao mercado, a nossa equipa está a analisar as suas informações com prioridade.",
    "Assim que a nossa triagem for concluída e a sua página interativa for gerada, receberá uma nova notificação oficial por aqui.",
    "TRIAGEM EM ANDAMENTO. RETORNO EM BREVE.",
    "rgba(0,191,255,0.07)",
    "#00bfff"
  );

  GmailApp.sendEmail(email, assunto, '', { htmlBody: corpoHtml, name: "Osorio Ai Labs | Social", from: EMAIL_REMETENTE });
}

// 4. E-mail de Aprovação
function enviarEmailAprovacao(nome, email) {
  var assunto = "Parabéns! Currículo Aprovado - Fazendo a Nossa Parte";
  var corpoHtml = gerarTemplateEmail(
    "Perfil Aprovado!",
    nome,
    "Temos excelentes notícias! O seu perfil foi analisado e <strong>aprovado</strong> pela nossa equipa de especialistas.",
    "A sua página web interativa já foi gerada e o seu currículo está oficialmente disponível na nossa Vitrine de Talentos.",
    "A partir de agora, empresas parceiras e recrutadores já podem aceder ao seu perfil premium e entrar em contacto direto consigo pelas vias que informou no formulário. Desejamos muito sucesso na sua jornada!",
    "O SEU PERFIL JÁ ESTÁ ONLINE NA NOSSA VITRINE.",
    "rgba(0,255,136,0.07)",
    "#00ff88"
  );

  GmailApp.sendEmail(email, assunto, '', { htmlBody: corpoHtml, name: "Osorio Ai Labs | Social", from: EMAIL_REMETENTE });
}

// 5. E-mail de Recusa
function enviarEmailRecusa(nome, email) {
  var assunto = "Atualização de Status: Fazendo a Nossa Parte";
  var corpoHtml = gerarTemplateEmail(
    "Atualização de Status",
    nome,
    "Agradecemos o seu interesse e o tempo dedicado a preencher o formulário da iniciativa <strong>Fazendo a Nossa Parte</strong>.",
    "Após uma análise cuidadosa da nossa equipa, informamos que o seu perfil não foi selecionado para avançar para a Vitrine de Talentos neste momento.",
    "O nosso espaço é limitado e procuramos manter um equilíbrio específico de vagas e setores. O seu currículo permanecerá na nossa base de dados interna para oportunidades futuras. Continue a acompanhar os nossos projetos!",
    "AGRADECEMOS A SUA PARTICIPAÇÃO.",
    "rgba(255,51,102,0.07)",
    "#ff3366"
  );

  GmailApp.sendEmail(email, assunto, '', { htmlBody: corpoHtml, name: "Osorio Ai Labs | Social", from: EMAIL_REMETENTE });
}

// ==========================================
// MOTOR DE TEMPLATES (Evita repetição de código)
// ==========================================
function gerarTemplateEmail(titulo, nome, p1, p2, p3, destaqueBox, bgDestaque, corDestaque) {
  return `
  <!DOCTYPE html>
  <html>
  <head><meta charset="UTF-8"></head>
  <body style="margin: 0; padding: 0; background-color: #0f0f14; font-family: 'Roboto', Arial, sans-serif; color: #e8e8f0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #0f0f14; padding: 40px 20px;">
      <tr><td align="center">
        <table width="600" cellpadding="0" cellspacing="0" border="0" style="background-color: #1a1a24; border: 1px solid rgba(0,191,255,0.12); border-radius: 12px; overflow: hidden;">
          <tr>
            <td style="padding: 40px 30px 30px 30px;">
              <h2 style="margin: 0 0 20px 0; font-family: 'Montserrat', Arial, sans-serif; font-size: 20px; color: ${corDestaque}; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">
                ${titulo}
              </h2>
              <p style="margin: 0 0 15px 0; font-size: 16px; line-height: 1.6; color: #e8e8f0;">Olá, <strong>${nome}</strong>.</p>
              <p style="margin: 0 0 15px 0; font-size: 15px; line-height: 1.6; color: #8888aa;">${p1}</p>
              <p style="margin: 0 0 15px 0; font-size: 15px; line-height: 1.6; color: #8888aa;">${p2}</p>
              <p style="margin: 0 0 30px 0; font-size: 15px; line-height: 1.6; color: #8888aa;">${p3}</p>
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center" style="background-color: ${bgDestaque}; border: 1px solid ${corDestaque}40; border-radius: 8px; padding: 15px;">
                    <p style="margin: 0; font-size: 14px; font-weight: 600; color: ${corDestaque}; letter-spacing: 1px; text-transform: uppercase;">
                      ${destaqueBox}
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td align="center" style="background-color: #1a1a24; padding: 0 30px 30px 30px;">
              <img src="https://raw.githubusercontent.com/Osorio-AI-Labs/Imagem-de-e-mail-Site-Osorio-AI-Lbas/main/Suporte-assinatura-osorio.jpg" width="540" style="display: block; max-width: 100%; height: auto; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
            </td>
          </tr>
          <tr>
            <td align="center" style="padding: 20px; background-color: #0f0f14; border-top: 1px solid rgba(0,191,255,0.12);">
              <p style="margin: 0; font-size: 12px; color: #555577;">© 2026 Osorio Ai Labs | osorioailabs.social.br</p>
            </td>
          </tr>
        </table>
      </td></tr>
    </table>
  </body>
  </html>
  `;
}

function testarEnvioAlias() {
  var meuEmail = "fabianotrabalhotrb@gmail.com"; // Vai enviar para si mesmo
  GmailApp.sendEmail(meuEmail, "Teste Direto Apps Script", "Se está a ler isto, o script já consegue usar o alias!", {
    from: "social@osorioailabs.com.br"
  });
}

// ==========================================
// 6. API DA VITRINE: Envia TODOS os dados dos aprovados (LEITURA POR ÍNDICES FIXOS)
// ==========================================
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dados = sheet.getDataRange().getValues();
  var aprovados = [];

  for (var i = 1; i < dados.length; i++) {
    var status = dados[i][1]; // Coluna B (Índice 1)

    // Verifica se está aprovado, blindando contra espaços vazios
    if (status && status.toString().trim() === "Aprovado") {
      aprovados.push({
        nome: dados[i][2] ? dados[i][2].toString().trim() : "",         // Coluna C
        email: dados[i][3] ? dados[i][3].toString().trim() : "",        // Coluna D
        whatsapp: dados[i][4] ? dados[i][4].toString().trim() : "",     // Coluna E
        cidade: dados[i][5] ? dados[i][5].toString().trim() : "",       // Coluna F
        linkedin: dados[i][6] ? dados[i][6].toString().trim() : "",     // Coluna G
        foto: dados[i][7] ? dados[i][7].toString().trim() : "",         // Coluna H
        cargo: dados[i][8] ? dados[i][8].toString().trim() : "",        // Coluna I
        experiencia: dados[i][9] ? dados[i][9].toString().trim() : "",  // Coluna J
        tags: dados[i][10] ? dados[i][10].toString().trim() : "",       // Coluna K
        resumo: dados[i][11] ? dados[i][11].toString().trim() : "",     // Coluna L
        historico: dados[i][12] ? dados[i][12].toString().trim() : "",  // Coluna M
        formacao: dados[i][13] ? dados[i][13].toString().trim() : "",   // Coluna N
        competencias: dados[i][14] ? dados[i][14].toString().trim() : ""// Coluna O
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify(aprovados))
    .setMimeType(ContentService.MimeType.JSON);
}
