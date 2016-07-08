/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
        /*function accessUser() {
           //var code = $_GET('access_token');
           alert('hello');
            var code = getToken();
            localStorage.setItem('accessToken', code)
            //$("#code").append('helo');
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
                  //document.getElementById("login").innerHTML = "";
                  //document.body.style.backgroundColor = "white";
                  var email = str.user.mail;
                  localStorage.setItem("email", email);
                  //window.location.href = "project-page.html"
                  //$("#project-page").append(projectPage);
                  //loadListProjects();
                  var userId = str.user.uid;
                  localStorage.setItem("uId", userId);
                  self.close();
              })
             .fail(function (jqXHR, textStatus, errorType) {
                 app.showNotification(textStatus + ' ' + errorType);
                 //myWindow.close();
                 //self.close();
             });



        };*/
    };

    return app;
})();