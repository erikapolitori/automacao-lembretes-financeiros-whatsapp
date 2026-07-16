const TOKEN = 'SEU_TOKEN_AQUI';
const PHONE_NUMBER_ID = 'SEU_PHONE_NUMBER_ID';
const PLANILHA_ID = 'SEU_ID_DA_PLANILHA';

const TEMPLATE_NAME = 'lembrente_financeiro';
const TEMPLATE_LANGUAGE = 'pt_BR';

function enviarLembretesDeContas() {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);

  const meses = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
  ];

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const hojeEhSegunda = hoje.getDay() === 1;
  const nomeAbaAtual = meses[hoje.getMonth()];
  const aba = planilha.getSheetByName(nomeAbaAtual);

  if (!aba) {
    throw new Error(`A aba "${nomeAbaAtual}" não foi encontrada.`);
  }

  const dados = aba.getDataRange().getValues();

  if (dados.length < 2) {
    Logger.log(`A aba ${nomeAbaAtual} não possui contas cadastradas.`);
    return;
  }

  const limiteSemana = new Date(hoje);
  limiteSemana.setDate(hoje.getDate() + 7);

  const enviosPorTelefone = {};

  for (let i = 1; i < dados.length; i++) {
    const diaVencimento = Number(dados[i][0]);
    const conta = String(dados[i][1] || '').trim();
    const valor = dados[i][2];
    const codigoPagamento = String(dados[i][3] || '').trim();

    const telefones = String(dados[i][4] || '')
      .split(';')
      .map(function(numero) {
        return numero.replace(/\D/g, '');
      })
      .filter(function(numero) {
        return numero;
      });

    const status = String(dados[i][5] || '').trim().toLowerCase();
    const enviadoHoje = String(dados[i][6] || '').trim().toLowerCase();
    const enviadoSemana = String(dados[i][7] || '').trim().toLowerCase();

    if (!diaVencimento || !conta || valor === '' || telefones.length === 0) {
      continue;
    }

    if (status !== 'pendente') {
      continue;
    }

    const vencimento = new Date(
      hoje.getFullYear(),
      hoje.getMonth(),
      diaVencimento
    );
    vencimento.setHours(0, 0, 0, 0);

    const venceHoje = vencimento.getTime() === hoje.getTime();
    const venceNaSemana = vencimento >= hoje && vencimento <= limiteSemana;

    const deveEnviarHoje = venceHoje && enviadoHoje !== 'sim';
    const deveEnviarSemana = hojeEhSegunda && venceNaSemana && enviadoSemana !== 'sim';

    if (!deveEnviarHoje && !deveEnviarSemana) {
      continue;
    }

    telefones.forEach(function(telefone) {
      if (!enviosPorTelefone[telefone]) {
        enviosPorTelefone[telefone] = [];
      }

      enviosPorTelefone[telefone].push({
        linha: i + 1,
        vencimento,
        conta,
        valor,
        codigoPagamento,
        deveEnviarHoje,
        deveEnviarSemana
      });
    });
  }

  const telefonesDestino = Object.keys(enviosPorTelefone);

  if (telefonesDestino.length === 0) {
    Logger.log('Nenhuma conta nova para enviar hoje.');
    return;
  }

  telefonesDestino.forEach(function(telefone) {
    const itens = enviosPorTelefone[telefone];

    const itensSemana = itens.filter(function(item) {
      return item.deveEnviarSemana;
    });

    const itensHojeComPagamento = itens.filter(function(item) {
      return item.deveEnviarHoje && item.codigoPagamento;
    });

    if (itensSemana.length > 0) {
      enviarWhatsApp(telefone, montarMensagemSemana(itensSemana));
    }

    if (itensHojeComPagamento.length > 0) {
      enviarWhatsApp(telefone, montarMensagemPagamentosHoje(itensHojeComPagamento));
    }
  });

  marcarLinhasComoEnviadas(aba, enviosPorTelefone);

  SpreadsheetApp.flush();
  Logger.log('Mensagens enviadas e planilha atualizada.');
}

function marcarLinhasComoEnviadas(aba, enviosPorTelefone) {
  const linhasMarcadasHoje = {};
  const linhasMarcadasSemana = {};

  Object.keys(enviosPorTelefone).forEach(function(telefone) {
    enviosPorTelefone[telefone].forEach(function(item) {
      if (item.deveEnviarHoje) {
        linhasMarcadasHoje[item.linha] = true;
      }

      if (item.deveEnviarSemana) {
        linhasMarcadasSemana[item.linha] = true;
      }
    });
  });

  Object.keys(linhasMarcadasHoje).forEach(function(linha) {
    aba.getRange(Number(linha), 7).setValue('sim');
  });

  Object.keys(linhasMarcadasSemana).forEach(function(linha) {
    aba.getRange(Number(linha), 8).setValue('sim');
  });
}

function montarMensagemSemana(itens) {
  const partes = itens.map(function(item) {
    return `${item.conta} - ${formatarData(item.vencimento)} - R$ ${formatarValor(item.valor)}`;
  });

  return limparTextoParaTemplate(
    `Contas da semana: ${partes.join(' | ')}`
  );
}

function montarMensagemPagamentosHoje(itens) {
  const partes = itens.map(function(item) {
    return `• ${item.conta} - R$ ${formatarValor(item.valor)} Pix/código: ${item.codigoPagamento}`;
  });

  return limparTextoParaTemplate(
    `⚠️ Vencem hoje: ${partes.join(' | ')}`
  );
}

function enviarWhatsApp(telefone, mensagem) {
  if (!telefone) {
    throw new Error('Telefone destino vazio. Preencha a coluna enviar_para.');
  }

  const url = `https://graph.facebook.com/v25.0/${PHONE_NUMBER_ID}/messages`;

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: telefone,
    type: 'template',
    template: {
      name: TEMPLATE_NAME,
      language: {
        code: TEMPLATE_LANGUAGE
      },
      components: [
        {
          type: 'body',
          parameters: [
            {
              type: 'text',
              parameter_name: 'resumo_contas',
              text: mensagem
            }
          ]
        }
      ]
    }
  };

  const opcoes = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${TOKEN}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  Logger.log(`Enviando para: ${telefone}`);
  Logger.log(`Mensagem: ${mensagem}`);

  const resposta = UrlFetchApp.fetch(url, opcoes);
  const codigoHttp = resposta.getResponseCode();
  const conteudo = resposta.getContentText();

  Logger.log(`HTTP ${codigoHttp}`);
  Logger.log(conteudo);

  if (codigoHttp < 200 || codigoHttp >= 300) {
    throw new Error(`A Meta recusou a mensagem. HTTP ${codigoHttp}: ${conteudo}`);
  }
}

function limparTextoParaTemplate(texto) {
  return String(texto || '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function formatarData(data) {
  return Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy');
}

function formatarValor(valor) {
  return Number(valor).toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
