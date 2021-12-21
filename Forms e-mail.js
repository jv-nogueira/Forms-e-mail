var mailApp=MailApp;
var app=SpreadsheetApp;
var spreadsheet=app.getActiveSpreadsheet();
var sheet=spreadsheet.getSheetByName("Formulario");


var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

var lastRow = sheet.getLastRow();

Utilities.sleep(30000);

function Email(){

  var values=sheet.getDataRange().getValues();
 
    for (var row=0; row < values.length; row++){
    if(row > lastRow-2) {
 
 SpreadsheetApp.flush()


var msg = "Olá "+values[row][1]+" =D\n\nEm resposta a solicitação do orçamento para o evento no dia "+(`00${values[row][5].getDate()}`).slice(-2)+"/"+(`00${values[row][5].getMonth() + 1}`).slice(-2)+"/"+values[row][5].getFullYear()+" das "+values[row][8].getHours()+":"+("00"+values[row][8].getMinutes()).slice(-2)+" até às "+values[row][9].getHours()+":"+("00"+values[row][9].getMinutes()).slice(-2)+" em "+values[row][6]+".\nProposta - Os atendentes vão chegar 2 horas antes para organizar e higienizar o local de trabalho e vão atender os convidados até o final do evento.\n\nO serviço da equipe terá o valor de "+values[1][16].toLocaleString('pt-br',{style: 'currency', currency: 'BRL'})+".\n\nEm caso de hora extra será cobrado um valor adicional de R$---, tendo de ser pagos conforme negociação.\n\n𝐂𝐨𝐧𝐭𝐫𝐚𝐭𝐨 𝐝𝐨 𝐬𝐞𝐫𝐯𝐢𝐜̧𝐨: por favor verifique o contrato de serviço, ao aceitar a cotação você esta concordando com os termos do contrato.\n𝐇𝐨𝐫𝐚 𝐞𝐱𝐭𝐫𝐚: Em caso de prolongar o horário do evento, o acréscimo dos valores é correspondente aos atendentes que permanecerão.\n𝐏𝐫𝐚𝐳𝐨 𝐝𝐚 𝐜𝐨𝐭𝐚𝐜̧𝐚̃𝐨: Esta cotação permanecerá válida durante 2 dias úteis a contar da data do envio deste e-mail. Seu caso está encerrado no suporte Servissorria, mas respondendo dentro desses 2 dias com o aceita da cotação, ele será reaberto. Você ainda terá também 10 dias para retornar, mas uma nova cotação será gerada. Depois de 10 dias não receberemos sua resposta de e-mail e um novo contato no suporte será necessário.\n𝐃𝐢𝐬𝐩𝐨𝐧𝐢𝐛𝐢𝐥𝐢𝐝𝐚𝐝𝐞 𝐝𝐞 𝐝𝐚𝐭𝐚𝐬: A Servissorria faz reserva prévia após o pagamento de no mínimo 40% do valor. A confirmação do agendamento ocorre após o pagamento e com a abertura do chamado. Caso alguma data não esteja disponível e impeça o agendamento do evento, será realizado um estorno integral do valor pago.\nPossuímos dois métodos de pagamento:\n    1. Em até 10x sem juros no cartão de crédito.\n    2. No boleto Bancário para ser pago em até 03 dias úteis.\nPreciso que informe se tiver interesse em seguir com o serviço, por favor responda a este email sem alterar o assunto com a opção de pagamento escolhida e as seguintes informações:\n\nForma de pagamento:\nNome Completo:\nCPF ou CNPJ:\nData de Nascimento:\nCEP:\nRua:\nNº:\nComplemento:\nBairro:\nCidade-Estado:\n\nOs cartões aceitos são: AMERICAN EXPRESS, VISA, ELO, MASTER CARD.\n𝐀𝐓𝐄𝐍𝐂̧𝐀̃𝐎: Não envie os dados do cartão por e-mail, caso tenha interesse nós entraremos em contato.\n*Os preços e condições desta proposta estão sujeitos a mudança até a data da emissão da nota fiscal, nas hipóteses de aumento de carga tributária, aumento de preços de insumos, variação da cotação TAX/BACEN do dólar norte americano vigente na data de emissão desta proposta em percentual superior a 3% para insumos, ou fatores fora do controle da Servissorria. Aplicam-se a esta venda os Termos e Condições Gerais de Venda da Servissorria.\n\nAguardaremos resposta. =D\n\nObrigado(a) por escolher a Servissorria."

mailApp.sendEmail(values[row][3], "Sobre "+values[row][4], msg);



    } 
    } 
} 

