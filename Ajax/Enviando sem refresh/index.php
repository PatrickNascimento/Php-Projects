<html>
  <head>
  <link rel="stylesheet" href="style.css">
    <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <script>
      $(function () {
        $('form').on('submit', function (e) {
          e.preventDefault();
          $.ajax({
            type: 'post',
            url: 'rota.php',
            data: $('form').serialize(),
            success: function () {
              alert('Dados enviados com sucesso!');
            }
          });
        });
      });
    </script>
  </head>
  <body>
  <div class="form-style-5">
    <form>
    <label ty>Nome</label>
      <input type="text"  name="nome" value="patrick"><br>
      <label>email</label>
      <input type="email" name="email" value="patrick"><br>
      <input name="submit" type="submit" value="Submit">
    </form>
    <div class="send" style="display: none;"><p>ENVIADO</p></div>    
    </div>
  </body>
</html>