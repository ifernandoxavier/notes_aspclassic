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
   sql = "SELECT * from tb_professor where cd_professor = "&cod
'  Response.Write(sql)
'  Response.End()
   set rs = conexao.execute(sql)
   nome = rs("ds_professor")
   cpf = rs("cpf")
   nascimento = rs ("data_nascimento")
   endereco = rs("endereco")
   cidade = rs("cidade")
   email = rs("email")
   curriculo = rs("curriculo")       

   bot        = "Alterar"
else
   bot       = "Incluir"
end if

%>

<div class="container">
<br><br>
  <form class="form-horizontal" action="manu_professor.asp" method="post">

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Nome:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="nome" placeholder="Nome" name="nome" value="<%=nome%>">
      </div>
    </div>

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>CPF:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="cpf" placeholder="Cpf" name="cpf" value="<%=cpf%>">
      </div>
    </div>

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Nascimento:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="nascimento" placeholder="Nascimento" name="nascimento" value="<%=nascimento%>">
      </div>
    </div>    

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Endereço:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="endereco" placeholder="Endereco" name="endereco" value="<%=endereco%>">
      </div>
    </div>    

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Cidade:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="cidade" placeholder="Cidade" name="cidade" value="<%=cidade%>">
      </div>
    </div>    

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Curriculo:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="curriculo" placeholder="Curriculo" name="curriculo" value="<%=curriculo%>">
      </div>
    </div>    

    <div class="form-group">
      <label class="control-label col-sm-2" for="nome"><b>Email:</b></label>
      <div class="col-sm-4">
        <input type="text" class="form-control" id="email" placeholder="Email" name="email" value="<%=email%>">
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