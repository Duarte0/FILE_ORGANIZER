1. Copie toda a pasta do programa para o computador onde ele sera utilizado.
2. Edite o arquivo config.env e informe a pasta de entrada, a pasta do Excel e a pasta de relatorios.
	O valor {DESKTOP} pode ser usado para criar a pasta de entrada na Area de Trabalho do usuario.
3. Edite o arquivo regras.xlsx com as empresas, rotas e palavras-chave que serao usadas na distribuicao.
4. Execute o arquivo DistribuidorArquivos.exe para iniciar o monitoramento automatico da pasta de entrada.
5. Verifique os relatorios gerados na pasta definida no config.env e acompanhe os logs diarios para validar o processamento.
6. Quando a rota nao informar {MES} ou {COMPETENCIA}, o sistema cria automaticamente a ultima pasta com o mes do arquivo, por exemplo 04 para 04.2026.