# Automação Salesforce Field Service

A execução padrão valida a planilha e prepara a fila local, sem abrir o
navegador nem acessar o Salesforce. A consulta ao Salesforce só ocorre com a
opção explícita `--consultar-salesforce`.

## Dados configurados

- Planilha: `C:\automação salesforce\Geoex_Metropolitana.xlsx`
- Aba: `Notas sem evidências`
- Coluna: `Nota`
- Saída: `C:\automação salesforce\Respostas`

## Preparar a fila

Instale a dependência e execute:

```powershell
python -m pip install -r requirements.txt
python automacao.py
```

Os caminhos podem ser substituídos por argumentos:

```powershell
python automacao.py --planilha "C:\caminho\notas.xlsx" --aba "Notas sem evidências" --coluna "Nota" --saida "C:\caminho\Respostas"
```

### Interface desktop

No Windows, abra a interface gráfica com:

```powershell
python interface.py
```

Na janela, selecione a planilha e a pasta de respostas; ajuste a aba, a coluna
da nota e o limite de consultas; depois escolha se deseja consultar o Salesforce
e/ou enviar pelo WhatsApp. O campo **Nome exato do grupo WhatsApp** pode ser
alterado antes de iniciar. Quando o envio estiver habilitado, a interface pede
confirmação mostrando o grupo escolhido. O andamento e os erros aparecem na
própria janela, e o botão **Abrir relatório** abre o CSV após sua criação.
O login do Salesforce e a leitura do QR Code do WhatsApp continuam sendo feitos
nas janelas dedicadas do Chrome.

Para gerar o executável Windows e colocá-lo na Área de Trabalho, instale as
dependências de empacotamento uma vez e execute o script do projeto:

```powershell
python -m pip install -r requirements-build.txt
.\build_windows.ps1
```

O executável será criado como `Automacao Salesforce.exe` na Área de Trabalho;
os scripts-fonte permanecem nesta pasta do projeto. O executável inclui o
Python e as bibliotecas do programa, mas o Chrome continua sendo necessário
para as automações do Salesforce e WhatsApp. Para gerar uma versão atualizada
após mudanças no código, execute `.\build_windows.ps1` novamente.

O programa cria uma pasta por nota com `descricao.txt` contendo Texto code
ação, Código de ação, Data encermto. e CenTrab respon. Também registra esses
campos, a coluna `UTD` da planilha, o status, o motivo, os arquivos baixados e
erros em `relatorio_processamento.csv`, que abre no Excel. Notas repetidas são
processadas uma única vez. A cada execução, as notas da planilha são comparadas
com o relatório: notas que ainda não existem são acrescentadas com status
`Pendente`, e todas as linhas e resultados anteriores são preservados mesmo se
uma nota deixar de constar na planilha. Descrições já existentes não são
sobrescritas durante a sincronização. Mantenha esse CSV fechado no Excel durante
consultas ao Salesforce, pois os resultados são gravados progressivamente.

## Consultar uma nota no Salesforce

O modo Salesforce é opt-in e consulta uma nota por execução por padrão. Feche
eventuais janelas já abertas pelo perfil dedicado antes de iniciar:

```powershell
python automacao.py --consultar-salesforce
```

O Chrome será aberto com um perfil dedicado em
`%LOCALAPPDATA%\SalesforceAutomation\ChromeProfile`, separado do seu perfil
pessoal. Na primeira execução, conclua manualmente o login e o MFA nessa
janela; a sessão fica salva nesse perfil local para execuções futuras. Se a
espera pelo login expirar, a janela permanecerá aberta. Conclua o login,
feche essa janela normalmente para salvar a sessão e rode o comando novamente.
O script estabiliza o campo de busca, pesquisa a nota e clica no resultado da
Ordem de Serviço que contém essa nota (não pressiona Enter). Na OS, abre a aba
`Relacionados`, localiza o card `Compromissos de serviços` e clica em cada
Compromisso de Serviço associado. Em cada registro de Compromisso, aciona
`Baixar arquivos`; se o Salesforce apresentar uma lista, seleciona as imagens
e clica em `Baixar`. A lista é aguardada até que as linhas de arquivos sejam
carregadas, inclusive dentro dos componentes Shadow DOM do Salesforce. O Chrome
salva na pasta da nota, sem abrir o seletor de destino do sistema. Quando o
Salesforce entrega um ZIP, a automação aguarda o download terminar, extrai todas
as imagens e remove o ZIP. Os nomes das imagens recebem o número da nota como
prefixo.

Notas sem imagens ficam marcadas como `Sem imagens`. Se houver várias imagens
no ZIP, todas as imagens reconhecidas são extraídas; outros tipos de arquivo
são ignorados. Falhas ao localizar a OS, abrir os Compromissos de Serviço ou
baixar os arquivos são registradas no relatório.

Se necessário, escolha outro diretório local para o perfil:

```powershell
python automacao.py --consultar-salesforce --perfil-chrome "C:\caminho\perfil-salesforce"
```

Para consultar mais notas, aumente o limite explicitamente:

```powershell
python automacao.py --consultar-salesforce --limite 10
```

O limite conta somente notas ainda não concluídas. Notas novas são priorizadas;
notas com status `Concluída` ou `Sem imagens` são ignoradas, enquanto notas
antigas pendentes, com erro ou parciais podem ser tentadas novamente. Cada
resultado é salvo no relatório assim que a nota termina, permitindo retomar a
fila sem reprocessar as notas já concluídas. Se não houver notas para processar,
o Chrome não será aberto.

Os downloads dependem do acesso do Chrome aos arquivos do Salesforce e podem
falhar se a rede ou a sessão da organização bloquear os arquivos. Erros são
registrados no relatório. A consulta ao Salesforce não envia mensagens ao
WhatsApp por padrão.

## Publicar no WhatsApp

O envio é uma opção separada. Para publicar as notas concluídas e, por último,
enviar um ZIP com todo o conteúdo de `Respostas`:

```powershell
python automacao.py --enviar-whatsapp
```

O Chrome usa um perfil separado em
`%LOCALAPPDATA%\SalesforceAutomation\WhatsAppProfile`. Na primeira execução,
leia manualmente o QR Code do WhatsApp Web. O programa procura o grupo pelo
nome exato informado (padrão: `Pendências Geoex - Metropolitana`) e pede confirmação no terminal
uma vez antes de iniciar os envios. Após a confirmação, as notas e o ZIP são
enviados automaticamente, sem confirmações individuais. Só são publicadas
notas com status `Concluída` e imagens; notas sem imagens, pendentes ou com
erro não são publicadas individualmente. Uma imagem é anexada pelo seletor de
arquivo de fotos já disponível na conversa, evitando depender da abertura do
menu de anexos; várias imagens são anexadas juntas pelo menu `Fotos e vídeos`
como fotos originais. A descrição combinada com o motivo de impedimento, quando
preenchido, é enviada como legenda na mesma mensagem. Se o WhatsApp não aceitar
o lote de imagens ou não oferecer o campo de legenda, o envio da nota é
interrompido; a descrição não é enviada separadamente.

Também é possível consultar o Salesforce e enviar no mesmo comando:

```powershell
python automacao.py --consultar-salesforce --limite 100 --enviar-whatsapp
```

Use `--grupo-whatsapp "Nome exato do grupo"` para escolher outro grupo ou
`--perfil-whatsapp "C:\caminho\perfil"` para informar outro perfil. O arquivo
`Respostas.zip` é criado ao lado da pasta `Respostas`, evitando incluí-lo dentro
de si mesmo. Um registro local fora da pasta guarda os envios confirmados:
conteúdo idêntico já enviado não é publicado novamente; notas ou arquivos
alterados podem ser enviados novamente. Para o ZIP, o envio só é confirmado
quando uma nova mensagem com o nome do arquivo aparece no grupo e tem um
indicador acessível de envio concluído. A automação aguarda até cinco minutos;
se não conseguir confirmar, informa as mensagens observadas, não marca o ZIP
como enviado e encerra o navegador. Estados antigos sem essa confirmação serão
tratados como não verificados e o ZIP será enviado novamente uma vez.
Após confirmar o envio do ZIP e salvar o estado, o navegador é encerrado
automaticamente e o comando termina.

O programa não envia credenciais nem acessa uma API externa do WhatsApp; a
sessão permanece no perfil local do Chrome. Confirme cuidadosamente o grupo e
o conteúdo antes de autorizar o envio.

## Validação

```powershell
python -m unittest -v
```

## Validação no Salesforce

O vídeo de referência mostra a busca, o modal com as imagens e o download
iniciado pelo navegador. Há mais de uma Ordem de Serviço nos resultados: a
automação clica na que contém a nota, abre a aba `Relacionados`, entra no
Compromisso de Serviço correspondente e baixa os arquivos desse registro. Os
testes automatizados cobrem a estabilização da busca, a navegação para
Compromissos relacionados, ZIP, modal e download direto. Valide primeiro com
`--limite 1`.
