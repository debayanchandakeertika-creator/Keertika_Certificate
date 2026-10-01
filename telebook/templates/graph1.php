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
    <td><a href=/roundsgraph?days={{prev}}> Prev Day </a></td>
    <td><a href=/roundsgraph> Today </a></td>
    <td><a href=/roundsgraph?days={{next}}> Next Day </a></td>
  </tr>

</table>
 <canvas id="bar_chart" width=100% height=25%></canvas>

    <script>

var bar_chart=new Chart(document.getElementById("bar_chart"), {

    type: 'horizontalBar',
    data: {
      labels: {{names | safe}},
      datasets: [
        {
          label: "Number of Beacons Covered",
          backgroundColor: 'grey',
          barPercentage: 0.5,
          barThickness: 3,
          maxBarThickness: 1,
          minBarLength: 2,
          data: {{data | safe}},

        }
      ],
    },
    options: {
      scaleBeginAtZero : true,
      responsive: true,
      legend: { display: false },
      title: {
        display: true,
        text: 'Patrolling Data',
      },

      scales: {
        xAxes: [{
                ticks:{beginAtZero: true}
        }],
        yAxes: [{
                stacked: true,
                grid:{offset: true},
        }]
        }

    }


});
document.getElementById("bar_chart").onclick = function(evt){
			var activePoints = bar_chart.getElementsAtEvent(evt);
			var firstPoint = activePoints[0];
			var label = bar_chart.data.labels[firstPoint._index];
			if (firstPoint !== undefined)
        var names={{names|safe}};
        var empids={{empids|safe}};
        var empid=empids[names.indexOf(label)];
        document.location = ("/roundsgraph?days="+{{days|safe}}+"&empid="+empid.toString());

		};

    </script>

<canvas id="bar-chart1" width=100% height=30%></canvas>

<script>

new Chart(document.getElementById("bar-chart1"), {
    type: 'bar',
    data: {
      labels: ['0 Am','1 Am','2 Am','3 Am','4 Am','5 Am','6 Am','7 Am','8 Am','9 Am','10 Am','11 Am','12 Pm','13 Pm','14 Pm','15 Pm','16 Pm','17 Pm','18 Pm','19 Pm','20 Pm','21 Pm','22 Pm','23 Pm'],
      datasets: [
        {
          label: "Number of Beacons Covered",
          backgroundColor: 'grey',
          data: {{hourly}}
        }
      ]
    },
    options: {
      responsive: true,
      legend: { display: false },
      title: {
        display: true,
        text: 'Patrolling Hourly Data'
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