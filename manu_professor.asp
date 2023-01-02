<br>
<!--#include file="wpg_cnx.asp"-->
<!--#include file="funcoes.asp"-->
<%
session.LCID = 1046
'RECEBENDO CAMPOS DO FORMULÁRIO
opc = request.querystring("opc")
botao =  request.form("botao")
cod =  request.form("hfcod")

nome =  request.form("nome")
cpf =  request.form("cpf")
email =  request.form("email")
nascimento =  request.form("nascimento")
endereco =  request.form("endereco")
cidade =  request.form("cidade")
curriculo =  request.form("curriculo")

if botao = "Incluir" then
' INCLUINDO OS DADOS RECEBIDOS NA TABELA
    sql = "INSERT INTO tb_professor (ds_professor, cpf, email, data_nascimento, endereco, cidade, curriculo)"
    sql = sql & " VALUES('"&nome&"','"&cpf&"', '"&email&"', '"&data_banco(nascimento)&"','"&endereco&"','"&cidade&"','"&curriculo&"')"
'response.write sql
'response.end
conexao.execute(sql)
%>
<script>
    alert ("Dados incluído com sucesso!")
    parent.location = "cst_professor.asp"
</script>
<%
elseif botao = "Alterar" then
' ALTERANDO OS DADOS RECEBIDOS NA TABELA

  sql = "UPDATE tb_professor SET"
  sql = sql & " ds_professor = '"&nome&"',"
  sql = sql & " cpf = '"&cpf&"',"
  sql = sql & " email = '"&email&"',"
  sql = sql & " data_nascimento = '"&data_banco(nascimento)&"',"
  sql = sql & " endereco = '"&endereco&"',"
  sql = sql & " cidade = '"&cidade&"',"
  sql = sql & " curriculo = '"&curriculo&"'"     
  sql = sql & " WHERE cd_professor = "&cod
'response.write sql
'response.end
  conexao.execute(sql)
%>
<script>
    alert("Dados alterado com sucesso!")
    parent.location = "cst_professor.asp"
</script>
<%
elseif opc <> "" then
' DELETANDO REGISTRO SELECIONADO
   cod = request.querystring("cod")
   sql = "DELETE FROM tb_professor WHERE cd_professor="&cod
'  response.write sql
'  response.end
   conexao.execute(sql)
%>
<script>
    alert("Linha excluída com sucesso!")
    parent.location = "cst_professor.asp"
</script>
<%
end if
%>