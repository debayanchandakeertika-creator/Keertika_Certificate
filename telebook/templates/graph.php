<!doctype html>
<html lang="en">
<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-giJF6kkoqNQ00vy+HMDP7azOuL0xtbfIcaT9wjKHr8RbDVddVHyTfAAsrekwKmP1" crossorigin="anonymous">

    <!-- Optional JavaScript -->
    <!-- jQuery first, then Popper.js, then Bootstrap JS -->
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.5.0/Chart.min.js"></script>

  <!-- <meta http-equiv="refresh" content="3"> -->
</head>
<body>

<nav class="navbar navbar-expand-lg navbar-dark bg-dark">
  <div class="container-fluid">
    <a class="navbar-brand" href="#">
    <img src="https://eventaa.com/timthumb.php?src=profile_pic/627_profilepic.jpg&width=165&height=165&zc=1" alt="" width="30" height="24" class="d-inline-block align-top">
    </a>
    <a class="navbar-brand" href="#">
    <img src="https://www.achillesresolute.com/images/index/achillesresolute.png" alt="" width="30" height="24" class="d-inline-block align-top">
    </a>
    <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNavDropdown" aria-controls="navbarNavDropdown" aria-expanded="false" aria-label="Toggle navigation">
      <span class="navbar-toggler-icon"></span>
    </button>
    <div class="collapse navbar-collapse" id="navbarNavDropdown">
      <ul class="navbar-nav">
        <li class="nav-item">
          <a class="nav-link active" aria-current="page" href="http://188.166.10.173:7878/zone">Today</a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="http://188.166.10.173:7878/zone?days=3">Last 3 days</a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="http://188.166.10.173:7878/zone?all=y">All</a>
        </li>
        <li class="nav-item dropdown">
          <a class="nav-link dropdown-toggle" href="#" id="navbarDropdownMenuLink" role="button" data-bs-toggle="dropdown" aria-expanded="false">
            Download links
          </a>
          <ul class="dropdown-menu" aria-labelledby="navbarDropdownMenuLink">
            <li><a class="dropdown-item" href="http://188.166.10.173:7878/downloadData">Todays Data</a></li>
            <li><a class="dropdown-item" href="http://188.166.10.173:7878/downloadData1">Last 3 days Data</a></li>
            <li><a class="dropdown-item" href="http://188.166.10.173:7878/downloadData?all=y">All Data</a></li>
          </ul>
        </li>
      </ul>
    </div>
  </div>
</nav>





{{date}} IST &nbsp; {{hh}}<br>
<table width="200">

  <tr>
    <td><a href=/get_graph?days={{prev}}> Prev Day </a></td>
    <td><a href=/get_graph> Today </a></td>
    <td><a href=/get_graph?days={{next}}> Next Day </a></td>
  </tr>

</table>
 <canvas id="bar-chart" width="300" height="150"></canvas>

    <script>

new Chart(document.getElementById("bar-chart"), {
    type: 'bar',
    data: {
      labels: {{time | safe}},
      datasets: [
        {
          label: "Number of Beacons Covered",
          backgroundColor: ["grey"],
          data: {{data}}
        }
      ]
    },
    options: {
      responsive: true,
      legend: { display: false },
      title: {
        display: true,
        text: 'Patrolling Data'
      },

      scales: {
            yAxes: [{
                ticks: {
                    beginAtZero:true
                }
            }]
        }

    }


});

    </script>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta1/dist/js/bootstrap.bundle.min.js" integrity="sha384-ygbV9kiqUc6oa4msXn9868pTtWMgiQaeYH7/t7LECLbyPA2x65Kgf80OJFdroafW" crossorigin="anonymous"></script>

</body>
</html>