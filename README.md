# Importação automática de rotas

Esta automação consulta a pasta pública do Google Drive e importa a planilha de rotas no Firebase. O computador pode permanecer desligado.

## Arquivos necessários no GitHub

- `.github/workflows/automatic-route-import.yml`
- `automation/route_importer.py`
- `automation/requirements.txt`

## Colunas da planilha

Obrigatórias:

- Data
- Rota
- Ordem
- Nome da Loja
- Endereço
- Bairro
- CEP
- Cidade
- Estado
- Congelados
- Refrigerados
- Mercearia
- Padaria
- Outros
- Total Geral
- Motorista
- Conferente

Opcionais:

- Ajudante
- Valor da Nota

Motorista, conferente e ajudante podem ser informados pelo nome completo ou e-mail cadastrados no aplicativo. O ajudante pode ficar vazio.

## Configuração segura

No repositório do GitHub, crie um segredo chamado `FIREBASE_SERVICE_ACCOUNT` e cole nele o conteúdo completo do arquivo JSON de uma conta de serviço do Firebase. Nunca envie esse JSON como arquivo público do repositório.

## Funcionamento

O GitHub verifica a programação a cada 15 minutos. O horário, os dias e o nome esperado do arquivo são definidos no painel web em **Importação automática**.

O arquivo da pasta do Drive precisa estar disponível para qualquer pessoa com o link. Quando houver mais de uma planilha `.xlsx`, informe no painel o nome exato da planilha que deve ser importada.
