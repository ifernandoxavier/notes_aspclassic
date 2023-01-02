<!--#include file="topo.asp"-->  
<!--#include file="wpg_cnx.asp"-->  
<%
  'fazendo a query e executando a query
  sql = "SELECT * FROM tb_professor ORDER BY ds_professor "
  set conecExec = conexao.execute(sql)
%>

<script>
  function Excluir(cod)
  {
    if(confirm("Confirma exclusão?"))
     {
  	   parent.location = "manu_professor.asp?opc=exc&cod=" + cod ;
     }
  } 
</script>
<br>
<div class="container">
  <form action="frm_professor.asp" method=post>
  <button type="submit" class="btn btn-primary align-center">Adicionar</button>

  <h4>Professor</h4>
  <div class="table-responsive">          
  <table class="table">
    <thead>
      <tr>
        <th>#</th>
        <th>Nome</th>
        <th>CPF</th>
        <th>Data de Nascimento</th>
        <th>E-mail</th>
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
        <a href="frm_professor.asp?evt=alt&cod=<%=conecExec("cd_professor")%>">
          <img src="imagens/alt.jpg">
        </a>		
		  </td>
      <td><%=conecExec("ds_professor")%></td>
      <td><%=conecExec("cpf")%></td>
      <td><%=conecExec("data_nascimento")%></td>
      <td><%=conecExec("email")%></td>            
      <td>
        <a href="javascript:Excluir(<%=conecExec("cd_professor")%>)"> 
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