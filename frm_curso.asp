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
   sql = "SELECT * from tb_curso where cd_curso = "&cod
'  Response.Write(sql)
'  Response.End()
   set rs = conexao.execute(sql)
   nome = rs("ds_curso")
   ativo = rs("ativo")
   resumo = rs ("resumo")
   professor = rs ("cd_professor")

   bot        = "Alterar"
else
   ativo = 0
   bot       = "Incluir"
end if

%>

<div class="container">
<br><br>
  <form class="form-horizontal" action="manu_curso.asp" method="post">

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Nome:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="nome" placeholder="Nome" name="nome" value="<%=nome%>">
      </div>
    </div>

    <div class="form-group">
      <label class="control-label col-sm-2" for="ativo"><b>Ativo:</b></label>
      <div class="col-sm-4">
      <select class="form-control" name="ativo">
        <%if cint(ativo) = 1 then %>
        <option value="1">Ativo</option>
        <option value="0">Inativo</option>
        <%else%>
        <option value="0">Inativo</option>
        <option value="1">Ativo</option>
        <%end if%>
        </select>
      </div>
    </div>
    <div class="form-group">
      <label class="control-label col-sm-2" for="resumo"><b>Resumo:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="resumo" placeholder="Resumo" name="resumo" value="<%=resumo%>">
      </div>
    </div>    
    <%
    sql = "SELECT * FROM tb_professor ORDER BY ds_professor"
    set rs = conexao.execute(sql)
    %>
    <div class="form-group">
      <label class="control-label col-sm-2" for="professor"><b>Professor:</b></label>
      <div class="col-sm-4">
        <select class="form-control" name="professor">
        <%do while not rs.eof %>
          <option value="<%=rs("cd_professor")%>"><%=rs("ds_professor")%> </option>
        <%
        rs.movenext
        loop
        %>
        </select>
       
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