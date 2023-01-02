<!--#include file="topo.asp"-->  
<!--#include file="wpg_cnx.asp"-->  
<%
  'fazendo a query e executando a query
  sql = "SELECT * FROM tb_curso a INNER JOIN tb_professor b ON a.cd_professor=b.cd_professor ORDER BY ds_curso"
  set conecExec = conexao.execute(sql)
%>

<script>
  function Excluir(cod)
  {
    if(confirm("Confirma exclusão?"))
     {
  	   parent.location = "manu_curso.asp?opc=exc&cod=" + cod ;
     }
  } 
</script>
<br>
<div class="container">
  <form action="frm_curso.asp" method=post>
  <button type="submit" class="btn btn-primary align-center">Adicionar</button>

  <h4>Curso</h4>
  <div class="table-responsive">          
  <table class="table">
    <thead>
      <tr>
        <th>#</th>
        <th>Nome do Curso</th>
        <th>Professor</th>
        <th>Resumo</th>
        <th>Ativo</th>
        <th>#</th>		
      </tr>
    </thead>

    <tbody>
<%
'AQUI COMEÇA O DO WHILE ONDE TRARÁ AS LINHAS DE ACORDO COM O SELECT FEITO - EOF = END OF FILE
DO WHILE NOT conecExec.eof 
%>
    <tr>
      <td>
        <a href="frm_curso.asp?evt=alt&cod=<%=conecExec("cd_curso")%>">
          <img src="imagens/alt.jpg">
        </a>		
		  </td>
      <td><%=conecExec("ds_curso")%></td>
      <td><%=conecExec("ds_professor")%></td>
      <td><%=conecExec("resumo")%></td>
      <td><%
          if cint(conecExec("ativo")) = 1 then
             response.write "<font color=blue>Ativo</font>"
          else
             response.write "<font color=red>Inativo</font>"        
          end if
          %></td>            
      <td>
        <a href="javascript:Excluir(<%=conecExec("cd_curso")%>)"> 
        <img src="imagens/excluir.jpg">
        </a>		
		  </td>
    </tr>
<%
'LOOP DEPOIS DA LINHA PARA QUE SE REPITA ENQUANTO TIVER REGISTROS NO SELECT FEITO
conecExec.movenext
loop
%>	  

            </tbody>
          </table>
        </div>
      </form>
    </div>
  </body>
</html>