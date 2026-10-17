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

O programa cria uma pasta por nota com `descricao.txt` contendo Texto code
ação, Código de ação, Data encermto. e CenTrab respon. Também registra esses
campos, a coluna `UTD` da planilha, o status, o motivo, os arquivos baixados e
erros em `relatorio_processamento.csv`, que abre no Excel. Notas repetidas são
processadas uma única vez. Resultados já registrados são preservados quando a
fila é gerada novamente.

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

Os downloads dependem do acesso do Chrome aos arquivos do Salesforce e podem
falhar se a rede ou a sessão da organização bloquear os arquivos. Erros são
registrados no relatório. O script não edita registros nem envia mensagens.

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
