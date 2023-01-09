<!--#include file="topo.asp"-->  
<!--#include file="wpg_cnx.asp"-->  

<%
  'transformando em utf-8'
  Response.ContentType = "text/html"
  Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
  Response.CodePage = 65001
  Response.CharSet = "UTF-8"

  'fazendo a query e executando a query
  sql = "select * from tb_modulo a inner join tb_curso b on a.cd_curso=b.cd_curso order by a.ds_curso" 

  set conecExec = conexao.execute(sql)
%>

<script>
  function Excluir(cod)
  {
    if(confirm("Confirma exclusão?"))
     {
  	   parent.location = "manu_modulo.asp?opc=exc&cod=" + cod ;
     }
  } 
</script>
<br>
<div class="container">
  <form action="frm_modulo.asp" method=post>
  <button type="submit" class="btn btn-primary align-center">Adicionar</button>

  <h4>Módulo</h4>
  <div class="table-responsive">          
  <table class="table">
    <thead>
      <tr>
        <th>#</th>
        <th>Módulo</th>
        <th>Curso</th>
        <th>Texto</th>
        <th>Anexo</th>        
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
        <a href="frm_modulo.asp?evt=alt&cod=<%=conecExec("cd_modulo")%>">
          <img src="imagens/alt.jpg">
        </a>		
		  </td>
      <td><%=conecExec("ds_modulo")%></td>
      <td><%=conecExec("ds_curso")%></td>
      <td><%=conecExec("texto")%></td>
      <td><%=conecExec("anexo")%></td>           
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