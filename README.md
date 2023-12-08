# Automação de Compilação e Envio de Dados

Este projeto é uma automação que compila várias bases de dados em uma única planilha do Excel e a envia automaticamente por e-mail.

## Como usar

1. Coloque todos os arquivos de dados que deseja compilar na pasta "bases".
2. Execute o script Python.
3. O script irá compilar todos os dados em uma única planilha do Excel chamada "Vendas.xlsx".
4. Em seguida, o script enviará automaticamente um e-mail com a planilha anexada.

## Dependências

Este projeto depende das seguintes bibliotecas Python:

- os
- pandas
- smtplib
- email.mime.multipart
- email.mime.text
- email.mime.application

Por favor, consulte o arquivo requirements.txt para as versões exatas dessas bibliotecas.
