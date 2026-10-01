<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <!-- FONTAWESOME -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.9.0/css/all.min.css" />
  <!-- CUSTOM CSS -->
  <link rel="stylesheet" href="../styles/style.css" />
  <title>Login &rarr; NIS IMS</title>
</head>
<body>

  <div class="centre_container">
    <div class="form form--advanced form--register">
      <h3><span>NIS</span>ISM</h3>
      <form action="#">
        <div class="form__item">
          <i class="fas fa-user"></i>
          <input type="text" placeholder='Enter username' />
        </div>
        <div class="form__item">
          <i class="fas fa-lock"></i>
          <input type="password" placeholder='Enter password' />
          <i id='form__item__showpass' class='fa fa-eye form__item__showpass'></i>
        </div>
        <div class="form__item">
          <i class="fas fa-envelope"></i>
          <input type="email" placeholder='Enter your email' />
        </div>
        <button type='submit'>LOGIN</button>
      </form>
      <h4>Forgot Password?</h4>
    </div>
  </div>

  <script src='../js/showPassword.js'></script>

</body>
</html>
