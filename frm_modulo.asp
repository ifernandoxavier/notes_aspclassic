<!--#include file="topo.asp"-->  
<!--#include file="wpg_cnx.asp"-->  

<%
'RECEBENDO CÓDIGO COMO PARAMETRO E EVT PARA SABER SE VAI INCLUIR OU ALTERAR
cod = Request.QueryString("cod")
evt  = Request.QueryString("evt")

if cod <> "" then
   cod = cint(cod)
end if

if ucase(evt) = "ALT" then
   sql = "SELECT * from tb_modulo where cd_modulo = "&cod
'  Response.Write(sql)
'  Response.End()
   set rs = conexao.execute(sql)
   nome = rs("nome")
   texto = rs("texto")
   anexo = rs ("anexo")       

   bot        = "Alterar"
else
   bot       = "Incluir"
end if

%>

<div class="container">
<br><br>
  <form class="form-horizontal" action="manu_modulo.asp" method="post">
    <%
    sql = "SELECT * FROM tb_curso ORDER BY ds_curso"
    set rs = conexao.execute(sql)
    %>
    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Nome do curso:</b></label>
      <div class="col-sm-4">
        <select class="form-control" name="nome">
        <%do while not rs.eof %>
          <option value="<%=rs("cd_curso")%>"><%=rs("ds_curso")%> </option>
        <%
        rs.movenext
        loop
        %>
        </select>
      </div>
    </div>

    <div class="form-group">
      <label class="control-label col-sm-2" for="texto"><b>Texto:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="texto" placeholder="Texto" name="texto" value="<%=texto%>">
      </div>
    </div>

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Anexo:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="anexo" placeholder="Anexo" name="anexo" value="<%=anexo%>">
      </div>
    </div>    
   
    <div class="form-group">        
      <div class="col-sm-offset-2 col-sm-10">
        <button type="submit" class="btn btn-default" name="Botao" value="<%=bot%>"><%=bot%></button>
      </div>
    </div>

	<input type="hidden" name="hfcod" value="<%=cod%>">	
  </form>
</div>

</body>
</html>