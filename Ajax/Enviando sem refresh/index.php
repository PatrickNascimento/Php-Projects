<html>
  <head>
      <style>
          .form {
            height: 300px;
            width: 30%;
            padding-left: 50px;
            padding-right: 50px;
            padding-top: 50px;
            padding-bottom: 50px;
            background-color: Gainsboro;
          }
      </Style>
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">  
    <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <script>    
    $( document ).ready(function() {
        $(".alert alert-success").hide();
    });      

      $(function () {
        $('form').on('submit', function (e) {
          e.preventDefault();
          $.ajax({
            type: 'post',
            url: 'rota.php',
            data: $('form').serialize(),
            success: function () {              
              $('#msg').show();
            }
          });
        });
      });
    </script>
  </head>
  <body>
  <div class="form">
    <form>
    <label ty>Nome</label>
      <input class="form-control" type="text"  name="nome" value=""><br>
      <label>email</label>
      <input class="form-control" type="email" name="email" value=""><br>
      <input  class="btn btn-primary" name="submit" type="submit" value="Submit">
    </form>    
  <div id="msg" class="alert alert-success" style="display :none">
    <strong>Successo!</strong> Dados Enviados com sucesso!
  </div>
    </div>
  </body>
</html>