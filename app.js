const fs = require("fs")
const venom = require("venom-bot")
const xlsx = require("xlsx")

let atendente = fs.readFileSync("./DadosEnvio/atendente.txt", "utf-8")
let textoMsg = fs.readFileSync("./DadosEnvio/mensagem.txt", "utf-8")

const tabelaDados = xlsx.readFile("./DadosEnvio/BancoDeDados.xlsx")
const primeiraAba = tabelaDados.SheetNames[0];
const planilha = tabelaDados.Sheets[primeiraAba];
const dados = xlsx.utils.sheet_to_json(planilha);
//console.log(dados);


venom.create({
  browserPathExecutable: "C:\Program Files\Google\Chrome\Application\chrome.exe",
  session: "session-name"
})
  .then(client => start(client))
  .catch(erro => console.log(erro))

const start = (client) => {

  const mensagem = montaMensagem(atendente, dados, textoMsg)

  for (let i = 0; i < mensagem.length; i++) {

    setTimeout(() => {

        client.sendText(mensagem[i][0], mensagem[i][1]) // Envio da mensagem 
        .then(() => {
          console.log(`Mensagem enviada para o numero ${i+2} da tabela. ✅ NOME: ${padronizaNome(dados[i].nome)} | Telefone: ${dados[i].telefone}`);
        })
        .catch(() => {
          console.log(`Mensagem enviada para o numero ${i+2} da tabela. ❌ Não consegui manda mensagem pro número informado, talvez não exista o telefone ${dados[i].telefone}`);
        })
    }, i * tempoDeEnvio(10)) // Tempo de envio entre mensagens

  }
}



///////////////////////////////////////////////////////////////NAO MEXER////////////////////////////////////////////////////////////////////////////////
function converteCnpj(numeroCnpj) {
  numeroCnpj = numeroCnpj.toString()
  while(numeroCnpj.length < 14) {
    numeroCnpj = "0" + numeroCnpj
  }

  let numeroCnpjcortado = numeroCnpj.toString().split('')

  let valor1 = numeroCnpjcortado[0] + numeroCnpjcortado[1]
  let valor2 = numeroCnpjcortado[2] + numeroCnpjcortado[3] + numeroCnpjcortado[4]
  let valor3 = numeroCnpjcortado[5] + numeroCnpjcortado[6] + numeroCnpjcortado[7]
  let valor4 = numeroCnpjcortado[8] + numeroCnpjcortado[9] + numeroCnpjcortado[10] + numeroCnpjcortado[11]
  let valor5 = numeroCnpjcortado[12] + numeroCnpjcortado[13]

  let numeroCnpjatt = `${valor1}.${valor2}.${valor3}/${valor4}-${valor5}`

  return numeroCnpjatt
}

function tiraNove(telefone) {

  let ddi = telefone.substring(0, 2)
  let ddd = telefone.substring(2, 4)
  let numeroTelefone = telefone.substring(telefone.length - 8)

  let telefoneFinal = ''
  if (ddd < 30) {
    telefoneFinal = ddi + ddd + "9" + numeroTelefone
  } else {
    telefoneFinal = ddi + ddd + numeroTelefone
  }

  return telefoneFinal
}

function tempoDeEnvio(segundos) {
  const tempo = segundos * 1000
  return tempo
}

function montaMensagem(atendente, dadosEnvio, mensagem) {
  let dadosEMsg = []
  
  for (let i = 0; i < dadosEnvio.length; i++) {
    let listaCliente = padronizaNome(dadosEnvio[i].nome)
    let listaCnpj = dadosEnvio[i].cnpj
    let listaTelefones = dadosEnvio[i].telefone
       
    let texto = `Olá *${listaCliente},* meu nome é ${atendente.trim()}, falo em nome da CNPJ Legal. Tudo bem?\n\nEsse contato é referente ao CNPJ *${converteCnpj(listaCnpj)}.*\n\n${mensagem}`
  
    const trataTelefone = tiraNove(listaTelefones.toString().trim())
    dadosEMsg.push([`${trataTelefone}@c.us`, texto])
  
  }
  console.log(dadosEMsg)
 
  return dadosEMsg
}

function padronizaNome(nomes) {
  let nomeBruto = nomes.replace(/\d+|\./gm, '')
  let nomesMinusculo = nomeBruto.trim().toLowerCase().split(" ")
  
  let nomeCompleto = ''
  let nomePadronizado = []
    for (let i = 0; i < nomesMinusculo.length; i++) {
      const dividiNome = nomesMinusculo[i].split('') // armazena o array do nome
  
      const letraInicio = dividiNome[0].toUpperCase()
      dividiNome.splice(0, 1, letraInicio)
      
      nomePadronizado.push(dividiNome.join('').trim())
    }
      nomeCompleto += nomePadronizado.join(' ').trim()

    return nomeCompleto
}