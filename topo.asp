<!DOCTYPE html>
<html lang="en">
<head>
  <title>Desenvolvimento Web</title>
  <meta charset="windows-1252">
  <meta name="viewport" content="width=device-width, initial-scale=1">

  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css">
  <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js"  crossorigin="anonymous"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js" crossorigin="anonymous"></script>
</head>
<body>
  <nav class="navbar navbar-expand-lg navbar-light bg-light">
    <a class="navbar-brand" href="#">SISTEMA WEB</a>
    <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#conteudoNavbarSuportado" aria-controls="conteudoNavbarSuportado" aria-expanded="false" aria-label="Alterna navega��o">
      <span class="navbar-toggler-icon"></span>
    </button>

    <div class="collapse navbar-collapse" id="conteudoNavbarSuportado">
      <ul class="navbar-nav mr-auto">
        <li class="nav-item active">
          <a class="nav-link" href="inicio.asp">Home <span class="sr-only">(p�gina atual)</span></a>
        </li>
        <li class="nav-item dropdown">
          <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
          Cadastros</a>
          <div class="dropdown-menu" aria-labelledby="navbarDropdown">
            <a class="dropdown-item" href="cst_professor.asp">Professor</a>
            <a class="dropdown-item" href="cst_curso.asp">Curso</a>
            <a class="dropdown-item" href="cst_modulo.asp">M�dulo</a>          
            <a class="dropdown-item" href="upload.asp">Upload Exemplo</a>
            <a class="dropdown-item" href="cst_modulo_up.asp">M�dulo Upload</a>                    
          </div>
        </li>
      
      </ul>


      <button class="btn btn-outline-success my-2 my-sm-0" type="submit">Pesquisar  </button>
    </div>
  </nav>
</body>
</html>