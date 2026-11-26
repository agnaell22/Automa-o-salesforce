# Automação Salesforce Field Service

A execução padrão valida a planilha e prepara a fila local, sem abrir o
navegador nem acessar o Salesforce. A consulta ao Salesforce só ocorre com a
opção explícita `--consultar-salesforce`.

## Dados configurados

No ambiente de desenvolvimento, os padrões atuais são:

- Planilha: `D:\automação salesforce\Geoex_Metropolitana.xlsx`
- Aba: `Notas sem evidência`
- Coluna: `Nota`
- Saída: `D:\automação salesforce\Respostas`

No executável instalado, os dados ficam por padrão em `Automacao Salesforce`,
dentro de Documentos, conforme a localização resolvida pelo Windows.

## Preparar a fila

Instale a dependência e execute:

```powershell
python -m pip install -r requirements.txt
python automacao.py
```

Os caminhos podem ser substituídos por argumentos:

```powershell
python automacao.py --planilha "C:\caminho\notas.xlsx" --aba "Notas sem evidência" --coluna "Nota" --saida "C:\caminho\Respostas"
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
Os contadores mostram quantas notas foram concluídas no Salesforce e quantas
mensagens com imagens tiveram envio confirmado no WhatsApp, além do que ainda
falta consultar/reprocessar ou enviar. Eles são reconstruídos a partir do
relatório e do estado de envios salvos; após uma interrupção, basta abrir a
interface novamente para consultar o progresso persistido.
O login do Salesforce e a leitura do QR Code do WhatsApp continuam sendo feitos
nas janelas dedicadas do Chrome.

Para gerar o executável instalador Windows, instale as dependências de
empacotamento no computador de desenvolvimento e execute:

```powershell
python -m pip install -r requirements-build.txt
.\build_windows.ps1
```

O arquivo `Automacao Salesforce Instalador.exe` gerado na Área de Trabalho pode ser
copiado para outro computador Windows. No primeiro início, ele instala uma
cópia do executável e os fontes Python em
`%LOCALAPPDATA%\AutomacaoSalesforce\app`, cria um atalho na Área de Trabalho e
abre a interface instalada. A instalação é por usuário e não requer escrita
na raiz do disco nem privilégios de administrador.

Os dados de trabalho usam a pasta **Documentos** resolvida pelo próprio Windows
(incluindo nomes localizados e redirecionamentos, como OneDrive), dentro de
`Automacao Salesforce`; a planilha não é embutida no pacote. Se ela não estiver
nessa pasta, selecione-a na interface.
O executável inclui Python, bibliotecas e os fontes, mas o Chrome continua
sendo necessário para as automações Salesforce e WhatsApp. Para atualizar,
gere o pacote novo e execute-o; ele atualiza a cópia instalada e o atalho.

O programa cria uma pasta por nota com `descricao.txt` contendo Texto code
ação, Código de ação, Data encermto. e CenTrab respon. Também registra esses
campos, a coluna `UTD` da planilha, o `Status usuário`, o status de processamento,
o `status_wfm` extraído do campo `Status` do Compromisso de Serviço e a
comparação dos dois status em `relatorio_processamento.csv`, que abre no Excel.
Sempre que a planilha mudar, os metadados e o `descricao.txt` são sincronizados
automaticamente para todas as notas existentes. O status WFM consultado no
Salesforce também é acrescentado à legenda publicada no WhatsApp.
Na comparação atual, `VRNE` confere com status WFM `Incompleto`; outros valores
são considerados iguais somente quando a escrita é equivalente. O relatório
mostra `Confere`, `Nao confere` ou `Nao identificado`. Notas repetidas são
processadas uma única vez. A cada execução, as notas da planilha são comparadas
com o relatório: notas que ainda não existem são acrescentadas com status
`Pendente`, e todas as linhas e resultados anteriores são preservados mesmo se
uma nota deixar de constar na planilha. Descrições já existentes não são
sobrescritas manualmente, mas os dados da planilha são sincronizados a cada
execução. Mantenha esse CSV fechado no Excel durante
consultas ao Salesforce, pois os resultados são gravados progressivamente.

## Consultar uma nota no Salesforce

O modo Salesforce é opt-in e consulta uma nota por execução por padrão. Feche
eventuais janelas antigas já abertas pelo perfil dedicado antes de iniciar:

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
Depois de cada nota, a automação fecha a janela de seleção de arquivos e a guia
ativa da Ordem de Serviço ou do Compromisso de Serviço, inclusive quando não há
imagens ou ocorre uma falha. Isso evita que registros e modais se acumulem e
interceptem a busca seguinte.
Para evitar fechar um registro válido durante o carregamento, a seleção da OS
é validada pelo número completo da nota na sugestão de pesquisa; o código não
rejeita a OS com base no campo `Nota` da página, que pode não corresponder ao
campo indexado pela pesquisa do Salesforce.

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
Se a janela do Chrome ou a sessão Selenium for encerrada durante a fila, a
automação salva o erro da nota atual e interrompe as consultas restantes em vez
de continuar tentando com uma sessão inválida.

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
nome exato informado (padrão atual:
`GRUPO DE PENDENCIAS - GEOEX - SALVADOR`) e pede confirmação no terminal uma
vez antes de iniciar os envios. Após a confirmação, as notas e o ZIP são
enviados automaticamente, sem confirmações individuais. Só são publicadas
notas com status `Concluída` e imagens; notas sem imagens, pendentes ou com
erro não são publicadas individualmente. Uma imagem é anexada pelo seletor de
arquivo de fotos já disponível na conversa, evitando depender da abertura do
menu de anexos; no máximo duas imagens são anexadas juntas pelo menu
`Fotos e vídeos` como fotos originais. Se a pasta da nota tiver mais imagens,
todas continuam incluídas no ZIP enviado ao final. A descrição combinada com
o motivo de impedimento, quando preenchido, é enviada como legenda na mesma
mensagem. Se o WhatsApp não aceitar o lote de imagens ou não oferecer o campo
de legenda, o envio da nota é interrompido; a descrição não é enviada
separadamente. Se uma tentativa não puder ser confirmada, a próxima execução
pede que você verifique o grupo e confirme que a mensagem não foi enviada
antes de tentar novamente, evitando duplicatas. O campo de legenda pode levar
até 30 segundos para aparecer.

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
tratados como não verificados e o ZIP será enviado novamente uma vez. As notas
também aguardam até cinco minutos por um novo balão com imagem, legenda e
confirmação de envio antes de serem marcadas como enviadas; esse intervalo
impede que o envio da próxima nota comece enquanto a anterior ainda está sendo
processada.
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

## Estado da validação e pendências (25-11-2026)

### Verificado nesta versão

- Execução real no Salesforce com `--limite 3`: as três notas consultadas
  terminaram como `Concluída`. O log mostrou uma Ordem de Serviço e um
  Compromisso de Serviço por nota, além dos nomes dos arquivos e da pasta de
  destino de cada nota. Foram baixadas uma imagem, uma imagem e três imagens,
  respectivamente. Isso confirma o fluxo para essa amostra pequena, não para
  filas grandes nem para registros com vários compromissos.
- O executável instalador foi construído em Windows 11 com Python 3.13.6 e
  PyInstaller 6.22.3, sem substituir os executáveis existentes. O arquivo está
  em `build\verification\dist\Automacao Salesforce Instalador.exe` e o
  conteúdo do pacote inclui os quatro módulos Python necessários. A compilação
  em si passou; ainda não foi validada a instalação e execução em outro
  computador.
- A suíte tem 108 testes. Nesta execução, 107 passaram e um falhou porque o
  teste espera o grupo WhatsApp antigo `Pendências Geoex - Metropolitana`,
  enquanto a configuração atual usa `GRUPO DE PENDENCIAS - GEOEX - SALVADOR`.
  É preciso decidir qual nome é o desejado e alinhar teste e configuração.

### Avaliar em testes posteriores

- Aumentar gradualmente o volume Salesforce e acompanhar lentidão, falhas de
  navegação, retomada após erro e notas com múltiplos Compromissos de Serviço.
- Ao investigar imagens repetidas, comparar os IDs da Ordem de Serviço e do
  Compromisso de Serviço exibidos no log, os arquivos destinados a cada nota e
  seus hashes. Nomes de arquivo diferentes, sozinhos, não provam que as imagens
  tenham conteúdos diferentes.
- Conferir em uma execução real os valores `status_wfm` e `comparacao_status`
  no relatório e validar a legenda e o destino antes de autorizar o WhatsApp.
- Instalar e iniciar o executável em outra máquina ou em outro perfil Windows,
  sem Python instalado. Verificar a criação/atualização da instalação e do
  atalho, o acesso à pasta Documentos (inclusive redirecionada pelo OneDrive),
  a seleção da planilha e o login/MFA no Chrome. O Chrome continua sendo uma
  dependência externa do executável.
- O build registrou avisos sobre imports opcionais `pycparser.lextab` e
  `pycparser.yacctab`, mas terminou e gerou o arquivo. Verificar esses avisos
  novamente apenas se o executável apresentar falha ao iniciar ou executar.
