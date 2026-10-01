{% include "header.php" %}
<!-- TABLES START -->

<script type="text/javascript">
  function fetcd() {
    var a=document.getElementById('eid').value;
    url="fetchd.php?email="+a;

    var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
      var dec = this.responseText;
   
        document.getElementById("bluff").innerHTML="<h4 class=\"chart__title\">"+dec+"</h4>";
       
    }
  };
  xhttp.open("GET", url, true);
  xhttp.send();

   //console.log(a);
  }

  function chk_uname() {
    var a=document.getElementById('uname').value;
    url="chkuser.php?email="+a;

    var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
      var dec = this.responseText;

      if (dec=='0') {

        document.getElementById("balu").innerHTML="<h4 class=\"chart__title\" style='color:#f00;'>Already Exists</h4>";
      }else{

        document.getElementById("balu").innerHTML="<h4 class=\"chart__title\">available</h4>";

      }
      
    }
  };
  xhttp.open("GET", url, true);
  xhttp.send();

   //console.log(a);
  }

 function chk_email() {
    var a=document.getElementById('email').value;
    url="chkemail.php?email="+a;
   var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
      var dec = this.responseText;
      if (dec=='0') {
        document.getElementById("bale").innerHTML='<h4 class="chart__title" style="color:#f00;">Already Exists</h4>';
      } else if(dec=='2'){
        document.getElementById("bale").innerHTML='<h4 class="chart__title" style="color:#f00;">Valid email is required</h4>';
      }else{
        document.getElementById("bale").innerHTML='<h4 class="chart__title">AvAiLable</h4>';

      }

      
    }
  };
  xhttp.open("GET", url, true);
  xhttp.send();
  }

</script>

<div class="table_container">

  <div class="form form--advanced">

    {{usererror|safe}}
    <h4>Add Employee</h4>

    <form action="" method="POST">{%csrf_token%}
      <div class="form_rme">
        <h5>Login Details</h5>
      </div>
      <div class="form__item">
        <i class="fas fa-user"></i>
        <input type="text" id="uname" placeholder='Employee Username' name="emp_uname" onkeyup="chk_uname()" /> 
      </div>
      
     <div id="balu"></div>

      <div class="form__item">
        <i class="fas fa-lock"></i>
        <input type="password" placeholder='Enter password' name="emp_pass"/>
        <i id='form__item__showpass' class='fa fa-eye form__item__showpass'></i>
      </div>

      <div class="form_rme">
        <h5>Employee Details</h5>
      </div>
      <div class="form__item">
        <i class="fas fa-user"></i>
        <input type="text" placeholder='First name' name="fname"/>
      </div>

      <div class="form__item">
        <i class="fas fa-user"></i>
        <input type="text" placeholder='Last name' name="lname" />
      </div>

      <div class="form__item">
        <i class="fas fa-user"></i>
        <input type="text" placeholder='Email id' name="email" id="email" onkeyup="chk_email();" />
        <div id="bale"></div>
      </div>

      <div class="form__item">
        <i class="fas fa-phone-alt"></i>
        <input type="tel" placeholder='Enter phone number' name="phone" />
      </div>

      <div class="form_rme">
        <h5>Position &amp; Reporting to</h5>
      </div>

      <div class="form__item">
        Select Position
        <select name="pos">
            <option value="" disabled="disabled">Select Position</option>
            <option value="1">Director</option>
            <option value="2">Sr. Manager</option>
            <option value="3">Manager</option>
            <option value="4">Supervisor</option>
            <option value="5">Executive</option>
        </select>
      </div>

      <div class="form__item">
        <i class="fas fa-user"></i>
        <input type="number" id='eid' placeholder='Employee Id of the Reporting Person' name="reid" onkeyup="fetcd();" />
      </div>
<div id="bluff"> </div>
      <button type='submit'>Submit</button>

    </form>
  </div>

</div>
<!-- TABLES END -->

{% include "footer.php" %}
