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
ativo =  request.form("ativo")
resumo =  request.form("resumo")
professor =  request.form("professor")

if botao = "Incluir" then
' INCLUINDO OS DADOS RECEBIDOS NA TABELA
    sql = "INSERT INTO tb_curso (ds_curso, ativo, resumo, cd_professor)"
    sql = sql & " VALUES('"&nome&"',"&ativo&", '"&resumo&"', '"&professor&"')"
'response.write sql
'response.end
conexao.execute(sql)
%>
<script>
    alert ("Dados incluído com sucesso!")
    parent.location = "cst_curso.asp"
</script>
<%
elseif botao = "Alterar" then
' ALTERANDO OS DADOS RECEBIDOS NA TABELA

  sql = "UPDATE tb_curso SET"
  sql = sql & " ds_curso = '"&nome&"',"
  sql = sql & " ativo = '"&ativo&"',"
  sql = sql & " resumo = '"&resumo&"',"
  sql = sql & " cd_professor = '"&professor&"'"     
  sql = sql & " where cd_curso = "&cod
'response.write sql
'response.end
  conexao.execute(sql)
%>
<script>
    alert("Dados alterado com sucesso!")
    parent.location = "cst_curso.asp"
</script>
<%
elseif opc <> "" then
' DELETANDO REGISTRO SELECIONADO
   cod = request.querystring("cod")
   sql = "DELETE FROM tb_curso WHERE cd_curso="&cod
'  response.write sql
'  response.end
   conexao.execute(sql)
%>
<script>
    alert("Linha excluída com sucesso!")
    parent.location = "cst_curso.asp"
</script>
<%
end if
%>