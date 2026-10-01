<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <!-- CUSTOM STYLES -->
  <link rel="stylesheet" href="http://188.166.10.173/styles/style.css" />
  <!-- FONTAWESOME -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.9.0/css/all.min.css" />
  <title>NIS IMS</title>
</head>
<body>

  <!-- NAVBAR -->
  <?php include './components/navbar.php'; ?>

  <div class="main-container">
    <!-- COMPONENT_SIDEBAR -->
    <?php include './components/component_sidebar.php'; ?>
    <!-- MAIN_CONTENT -->
    <div class="main_content">
      <!-- CONTENT_NAV STARTS -->
      <div class="content_nav">
        <div class="content_nav__title">
          TABLES
        </div>
        <div class="content_nav__breadcrumbs">
          <a href="">home</a> / <a href="">tables</a>
        </div>
      </div>
      <!-- CONTENT_NAV ENDS -->

      <!-- TABLES START -->
      <div class="table_container">

        <div class="table table1">
          <h4>Basic Table</h4>
          <table>
            <tr>
              <th>#</th>
              <th>Heading1</th>
              <th>Heading2</th>
              <th>Heading3</th>
              <th>Heading4</th>
              <th>Date</th>
              <th>Action</th>
            </tr>
            <tr>
              <td>1</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>01-07-2019</td>
              <td><button>action</button></td>
            </tr>
            <tr>
              <td>2</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>01-07-2019</td>
              <td><button>action</button></td>
            </tr>
            <tr>
              <td>3</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>01-07-2019</td>
              <td><button>action</button></td>
            </tr>
            <tr>
              <td>4</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>Some title</td>
              <td>01-07-2019</td>
              <td><button>action</button></td>
            </tr>
          </table>
          <div class="table__footer">
            <div class="table__footer__entries">
              Showing 1 to 4 of 10 entries
            </div>
            <div class="table__footer__pagination">
              <button class='table__footer__pagination__prev'>
                <i class="fas fa-angle-left"></i>
              </button>
              <div class="table__footer__pagination__pages">
                <button>1</button>
                <button>2</button>
                <button>3</button>
              </div>
              <button class='table__footer__pagination__next'>
                <i class="fas fa-angle-right"></i>
              </button>
            </div>
          </div>
        </div>

      </div>
      <!-- TABLES END -->
    </div>
  </div>

  <script src='http://188.166.10.173/js/main.js'></script>

</body>
</html>
