# Objetivo

Esse programa tem como o objetivo a conversão para o padrão SHA256 de valores presentes em uma planilha

# Como Usar

Primeira o programa foi definido para funcionar ciretamente na pasta C:\User\Conversao
Ele busca um arquivo do tipo xlsx denominada Arquivo, portanto Arquivo.xlsx.

Ela espera que a primeira coluna esteja preenchida com os valores originais, sendo do tipo texto.

Qualquer coisa diferente dessa, vai resultar em erros, não concluído a tarefa.

# Como funciona

A ideia é que o programa vai ler todas as células manipuladas da primeira coluna, verificando se existe ou não algum valor a ser convertido.
Encontrando o valor, ele converte cria uma célula na coluna ao lado, inserindo o novo valor inserido.

Após concluir essa operação, ele cria o arquivo chamado Convetido.xlsx, e abre o arquivo solicitado.

# Importante

O arquivo sobrescreve o convertido automaticamente.

## Documentos consultados

### Trabalhar com arquivos

https://www.guj.com.br/t/descobrir-caminho-da-aplicacao/54712/5
https://www.guj.com.br/t/abrir-qualquer-arquivo-pelo-java/76178/11
https://www.devmedia.com.br/java-arquivos-e-fluxos-de-dados/22859
https://pt.stackoverflow.com/questions/319679/como-salvar-arquivo-de-uma-pasta-em-outra

### Trabalhar com planilhas

https://www.devmedia.com.br/apache-poi-manipulando-documentos-em-java/31778
https://www.baeldung.com/java-microsoft-excel

### Converter valores

https://www.baeldung.com/sha-256-hashing-java

### Tratar erros

https://www.w3schools.com/java/java_try_catch.asp

### Gerar executável

https://www.treinaweb.com.br/blog/criando-um-executavel-para-uma-aplicacao-java