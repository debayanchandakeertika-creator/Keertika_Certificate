{% include "header.php" %}
{%  load test_tag %}
      <div class="content_nav">
        <div class="content_nav__title">
          {{title_top}}
        </div>
        
      </div>
      
        <div class="table table1" style="margin: 1%;width: 98%;">
          <h4>{{content_tite}}</h4>

          <div class="form_container">

<table>
      <tr>
        <th>id</th>
        <th>Client Name</th>
        <th>Site Code</th>
        <th>Employee ID</th>
        <th>Issue Type</th>
        <th>Status</th>
        <th>Date</th>
        <th>Action</th>
      </tr>
{%for i in new23%}
      <tr>
        <td>{{i.id}}</td>
        <td>{{i.cid|cli_name}}</td>
        <td>{{i.cid|site_code}}</td>
        <td>{{i.eid}}</td>
        <td>{{i.type_of_issue}}</td>
        <td>{{i.status}}</td>
        <td>{{i.date_added}}</td>
        
        <td><button>action</button></td>
      </tr>
{% endfor%}
    </table>
  
  </div>

</div>

        </div>
      

{% include "footer.php" %}