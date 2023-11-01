# AtualizandoQueryExcel
 Uma solução para atualizar a consulta de um arquivo em excel que tem conexão com outra fonte de dados.
 
 Nosso fluxo de dados está no DataFlow do PowerBi, já que somente o proprietário do do fluxo pode fazer consultas e alterações a unica forma de atualizar os dados recentes é pela fonte de dados do excel.
 
 Para deixar o fluxo de atualização automatico e incremantar no script já existente eu precisava de alguma forma de pegar essa fonte de forma automatica, então achei essa solução na internet depois de muita pesquisa.

 Esse trecho do script consigo entrar no arquivo de excel atualizar os dados do dia e depois receber em um DataFrame do pandas para continuar o ETL.
