
    var qualityCriteriaId;
    var projectId;
    var ProductDescriptionId;
    var host = 'https://uprincecoredevapi.azurewebsites.net';
    var projectPage = '<div class="main-wrapper"> <header class="col-lg-12 col-md-12 col-sm-12 col-xs-12 header-top"> <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 no-padding full-height"> <div class="header-sub header-glyph full-height"> <p title="UPrince.Projects"> <span class="glyphicon glyphicon-folder-open" aria-hidden="true"></span> </div> <div class="header-sub h1-div"> <h1 class="roboto-light">Projects</h1> </div> <div class="header-sub" style="position:absolute;right:15px"><p class="fake-link" id="logOut" style="font-size:12px;font-weight: 100; vertical-align: middle"> Log Out</p> </div></div> </header> <section class="col-lg-12 col-md-12 col-sm-12 col-xs-12 modal-div relationship container no-padding"><div><input id="projectSearch"></div> <div id="listProjects" class="nav nav-pills nav-stacked"></div> </section>  </div>'
    var myWindow;
    var previous = 0;
    // The initialize function must be run each time a new page is loaded

    $(document).ready(function () {
        localStorage.setItem("loggedIn", 'false');

        accessUser();
    });

    function accessUser() {
        //var code = $_GET('access_token');
        var code = getToken();
        localStorage.setItem('accessToken', code)
        $("#status").append('');
        var url = "https://uprince-dev.pronovix.net/api/system/connect"
        var authorization = "Bearer " + code;
        //JQuery
        $.ajax({
            type: "POST",
            url: url,
            dataType: "json",
            //contentType: "application/json; charset=utf-8",
            headers: { "Authorization": authorization }
        })
          .done(function (str) {
             var email = str.user.mail;
              localStorage.setItem("email", email);
              //window.location.href = "project-page.html"
              var userId = str.user.uid;
              localStorage.setItem("uId", userId);
              localStorage.setItem('loggedIn', 'true');
              document.getElementById("test").innerHTML = localStorage.getItem('loggedIn');;

              //self.close();
              alert('done');
          })
         .fail(function (jqXHR, textStatus, errorType) {
             //app.showNotification(textStatus + ' ' + errorType);
             //myWindow.close();
             //self.close();
         });
    };


    function getToken() {
        var url = window.location.href;
        var startParam = url.indexOf('access_token');
        var start = url.indexOf('=', startParam) + 1;
        var eind = url.indexOf('&', start);
        return url.substring(start, eind);
    };
