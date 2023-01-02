<%
'definindo o ID de localidade do ASP
session.lcid = 1046

'usado para fazer a conexão ao banco de dados SQL
Set Conexao = Server.CreateObject("ADODB.Connection")

'Usado para armazenar um conjunto de registros de uma tabela de banco de dados
Set rec=Server.CreateObject("ADODB.recordset")

'String de conexão
Conexao.Open "Provider=MSOLEDBSQL;data source=localhost\sqlexpress;initial catalog=db_udasp;uid=xavier;pwd=1234"
%>