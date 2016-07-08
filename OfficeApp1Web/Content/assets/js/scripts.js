(function () {
    "use strict";
    var qualityCriteriaId;
    var projectId;
    var ProductDescriptionId;
    var ProjectName;
    var email = 'kurt@uprince.com';
    var host = 'http://uprincecoredevapi.azurewebsites.net';
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $("#openFile").click(function () {
                var x = document.getElementById("ProductID");
                projectId = x.elements[0].value;
                getProductDescription();
            });
            $("#saveJSON").click(saveJson);
            $(".btn").on('click', function () {
                var x = document.getElementById("email");
                email = x.elements[0].value;
                //loadListProjects();
                app.showNotification('hello');
                window.location.href = "project-menu.html";
            });
            $("#listProjects").on('click', 'li', function () {
                projectId = $(this).attr('id');
                window.location.href = "product-description-page.html";
                loadListProductDescriptions();
            });
            $("#listProductDescriptions").on('click', 'li', function () {
                ProductDescriptionId = $(this).attr('id');
                getProductDescription();
            });
            $("#prod").on('click', function () {
                host = 'http://uprincecoreprodapi.azurewebsites.net';
                loadListProjects();
            });
            $("#dev").on('click', function () {
                host = 'http://uprincecoredevapi.azurewebsites.net';
                loadListProjects();
            });
            loadListProjects();
            /*
                Fullscreen background
            */
            $.backstretch("/Content/assets/img/backgrounds/1.jpg");

            /*
                Form validation
            */
            $('.login-form input[type="text"], .login-form input[type="password"], .login-form textarea').on('focus', function () {
                $(this).removeClass('input-error');
            });

            $('.login-form').on('submit', function (e) {

                $(this).find('input[type="text"], input[type="password"], textarea').each(function () {
                    if ($(this).val() == "") {
                        e.preventDefault();
                        $(this).addClass('input-error');
                    }
                    else {
                        $(this).removeClass('input-error');
                    }
                });

            });


        });
        loadListProjects();
    }

    //load list with Projects, after getting the email adress
    function loadListProjects() {
        // $("#listProjects").html('');
        // $("#listProductDescriptions").html('');
        var dataEmail = {
            "customer": "",
            "email": email,
            "isFocused": {
                "customer": false,
                "title": false
            },
            "isRecycled": false,
            "orderField": "id",
            "sortOrder": "ASC",
            "status": {
                "Active": false,
                "All": true,
                "Closed": false,
                "New": false
            },
            "title": "",
            "toleranceStatus": {
                "All": true,
                "OutofTolerance": false,
                "Tolerancelimit": false,
                "WithinTolerance": false
            }
        };
        var data = { "title": "" }
        $.ajax({
            type: "POST",
            url: host + "/api/project/GetProjectList",
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
        .done(function (str) {
            var test = str;
            var length = Object.keys(str).length;
            for (var i = 0; i < length; i++) {
                var dummy = "<li id='".concat(str[i].id, "'><a href='#' class='active'>", str[i].title, "</a></li>");
                $("#listProjects").append(dummy);
            }
        })
    };


    //load list with PD, after getting the projectId
    function loadListProductDescriptions() {
        $("#listProductDescriptions").html('');
        var urlProject = host + '/api/ProductDescription/GetAllProductDescription?projectId=' + projectId;

        $.ajax({
            type: 'GET',
            url: urlProject,
            dataType: "json",
            jsonp: false,
            xhrFields: {
                withCredentials: false
            }
        })
            .done(function (str) {
                var length = Object.keys(str).length;
                $("#listProductDescriptions").append("<ul id = 'listProductDescriptions' class = 'projectList', style='list-style-type:circle'>");
                for (var i = 0; i < length; i++) {
                    var dummy = "<li id='".concat(str[i].Id, "'><a href='#' class='active'>", str[i].Title, "</a></li>");
                    //app.showNotification(dummy);
                    $("#listProductDescriptions").append(dummy)
                }
                $("#listProductDescriptions").append("</ul>")
            });
    };

    //get the Product Description Id as an String
    function getProductDescriptionUrl() {
        return host + '/api/productdescription?id=' + ProductDescriptionId;
    };

    //uses ajax to get JSON file with productdescription
    function getProductDescription() {
        var urlid = getProductDescriptionUrl();
        $.ajax({
            type: 'GET',
            url: urlid,
            dataType: "json",
            jsonp: false,
            xhrFields: {
                withCredentials: false
            }
        }).done(function (str) {

            var layout = "<head><style>table {border-collapse: collapse;} table, th, td {border: 1px solid black;text-align: 'left';  font-family: 'Calibri', 'sans-serif'} p,ol,ul{ font-family: 'Calibri', 'sans-serif'}</style></head>";
            var header = "<p class=MsoTitle>Product Description: ".concat(str.Title, "</p>");

            //purpose
            var purpose = "<h1>Purpose</h1>";
            var purposeJson = str.Purpose;
            //composition
            var composition = "<h1>Composition</h1>";
            var compositionJson = str.Composition;
            //derivation
            var derivation = "<h1>Derivation</h1>";
            var derivationJson = str.Derivation;
            //format
            var format = "<h1>Format and Presentation</h1>";
            var formatJson = str.FormatPresentation;
            var devSkills = "<h1>Development Skills Required</h1>"
            var devSkillsJson = str.DevSkills;

            //skills
            var skills = "<h1>Quality Criteria</h1>";
            var div = $("<div>")
                    .append(layout, header, purpose, purposeJson, composition, compositionJson, derivation,
                    derivationJson, format, formatJson, devSkills, devSkillsJson, skills);

            //tableQuality
            var tableSkills = "<table style='width:100%' id='table5' > <tr> <th>Quality Criteria </th> <th>Quality Tolerance </th> <th>Quality Method </th> <th>Quality Skills Required </th> </tr>";
            //checks how many criterias there are, write them out in table form
            qualityCriteriaId = [str.QualityCriteria.length];
            for (var i = 0; i < str.QualityCriteria.length; i++) {
                tableSkills = tableSkills.concat("<tr><td>".concat(str.QualityCriteria[i].Criteria,
                    "</td><td>", str.QualityCriteria[i].Tolerance, "</td><td>", str.QualityCriteria[i].Method,
                    "</td><td>", str.QualityCriteria[i].Skills).concat("</td></tr>"));
                qualityCriteriaId[i] = str.QualityCriteria[i].QualityCriteriaId;

            }

            //tableSkills = tableSkills.concat("<tr><td></td><td></td><td></td><td></td></tr>")  extra row 24 augustus
            div.append(tableSkills.concat("</table>"));

            //Qualityresponsibility
            var responsibilities = "<h1>Quality Responsibilities</h1>"
            if (str.QualityResponsibility === null) {
                var tableResponsibility = "<table style='width:100%'> <tr><th>Role</th><th>Responsible Individuals</th></tr><tr><th>Product Producer</th><td><p>"
                   .concat("", "</p></td></tr><tr><th>Product Reviewer(s)</th><td>",
               "",
               "</td></tr><tr><th>Product Approver(s)</th><td>",
               "", "</td></tr>");
                tableResponsibility = tableResponsibility.concat("</table>");
            }
            else {
                var tableResponsibility = "<table style='width:100%'> <tr><th>Role</th><th>Responsible Individuals</th></tr><tr><th>Product Producer</th><td><p>"
                    .concat(str.QualityResponsibility.Producer, "</p></td></tr><tr><th>Product Reviewer(s)</th><td>",
                str.QualityResponsibility.Reviewers,
                "</td></tr><tr><th>Product Approver(s)</th><td>",
                str.QualityResponsibility.Approvers, "</td></tr>");
                tableResponsibility = tableResponsibility.concat("</table>");
            }
            div.append(responsibilities, tableResponsibility);

            // insert HTML into Word document
            Office.context.document.setSelectedDataAsync(div.html(), { coercionType: "html" }, testForSuccess);
            /*var div = $("<div>")
                   .append(str.Title);
            Office.context.document.setSelectedDataAsync(div.html(), { coercionType: "html" }, testForSuccess)*/
        })

            .error(function (jqXHR, textStatus, errorThrown) {
                app.showNotification('fail')
            });

    };

    //adapt the JSON file with the latest info
    function saveJson() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Html,
             function (result) {
                 if (result.status === Office.AsyncResultStatus.Succeeded) {
                     var ID = ProductDescriptionId;
                     var html = result.value;
                     html = html.replace(/\s\s+/g, ' ');
                     //still has to clean out, the differents vars, was for testing
                     var title = extractTitle(html);
                     var purpose = extractChapter(html, ">Purpose", ">Composition");
                     $.ajax({
                         type: "POST",
                         url: host + "/api/productdescription/PostRTValue",
                         dataType: "json",
                         data: {
                             "ProductDescriptionIndex": "1",
                             "ProductDescriptionId": ID,
                             "ProductDescriptionProperty": purpose
                         },
                         success: function () {
                             app.showNotification('Success');

                         },

                         error: function () {
                             app.showNotification('Error');
                         }
                     });

                     var composition = extractChapter(html, ">Composition", ">Derivation");
                     $.ajax({
                         type: "POST",
                         url: host + "/api/productdescription/PostRTValue",
                         dataType: "json",
                         data: {
                             "ProductDescriptionIndex": "2",
                             "ProductDescriptionId": ID,
                             "ProductDescriptionProperty": composition
                         },
                         success: function () {
                             app.showNotification('Success');

                         },
                         error: function () {
                             app.showNotification('Error');
                         }
                     });

                     var derivation = extractChapter(html, ">Derivation", ">Format and Presentation");
                     $.ajax({
                         type: "POST",
                         url: demo + "/api/productdescription/PostRTValue",
                         dataType: "json",
                         data: {
                             "ProductDescriptionIndex": "3",
                             "ProductDescriptionId": ID,
                             "ProductDescriptionProperty": derivation
                         },
                         success: function () {
                             app.showNotification('Success');

                         },
                         error: function () {
                             app.showNotification('Error');
                         }
                     });

                     var formatPresentation = extractChapter(html, ">Format and Presentation", ">Development Skills Required");
                     $.ajax({
                         type: "POST",
                         url: host + "/api/productdescription/PostRTValue",
                         dataType: "json",
                         data: {
                             "ProductDescriptionIndex": "4",
                             "ProductDescriptionId": ID,
                             "ProductDescriptionProperty": formatPresentation
                         },
                         success: function () {
                             app.showNotification('Success');

                         },
                         error: function () {
                             app.showNotification('Error');
                         }
                     });

                     var devSkills = extractChapter(html, ">Development Skills Required", ">Quality Criteria");
                     $.ajax({
                         type: "POST",
                         url: host + "/api/productdescription/PostRTValue",
                         dataType: "json",
                         data: {
                             "ProductDescriptionIndex": "5",
                             "ProductDescriptionId": ID,
                             "ProductDescriptionProperty": devSkills
                         },
                         success: function () {
                             app.showNotification('Success');

                         },
                         error: function () {
                             app.showNotification('Error');
                         }
                     });
                     var qualityCriteria = extractDevSkills(html);
                     for (var i = 0; i < qualityCriteriaId.length; i++) {
                         var urlId = host + "/api/productdescription/DeleteQualityCriteria?criteriaId=" + qualityCriteriaId[i];
                         $.ajax({
                             type: "DELETE",
                             url: urlId,
                             success: function () {
                                 app.showNotification('Success');

                             },
                             error: function () {
                                 app.showNotification('Error');
                             }
                         });
                     }
                     for (var i = 0; i < qualityCriteria.length; i++) {
                         /*  var id;
                           if (i < qualityCriteriaId.length) {
                               id = qualityCriteriaId[i];
                           }
                           else{ id = null}*/
                         $.ajax({
                             type: "POST",
                             url: host + "/api/productdescription/PostQualityCriteria",
                             dataType: "json",
                             data: {
                                 "QualityCriteriaId": null,
                                 "ProductDescriptionId": ID,
                                 "Criteria": qualityCriteria[i][0],
                                 "Tolerance": qualityCriteria[i][1],
                                 "Method": qualityCriteria[i][2],
                                 "Skills": qualityCriteria[i][3]
                             },
                             success: function () {
                                 app.showNotification('Success');

                             },
                             error: function () {
                                 app.showNotification('Error');
                             }
                         });

                     }

                     var responsibilities = extractResponsibilities(html);
                     $.ajax({
                         type: "POST",
                         url: host + "/api/productdescription/PostQualityResponsibility",
                         dataType: "json",
                         data: {
                             "ProductDescriptionId": ID,
                             "Producer": responsibilities[0],
                             "Reviewer": responsibilities[1],
                             "Approver": responsibilities[2]
                         },
                         success: function () {
                             app.showNotification('Success');

                         },
                         error: function () {
                             app.showNotification('Error');
                         }
                     });

                     //app.showNotification(devSkills);
                     //app.showNotification(responsibilities[0]);
                 } else {
                     app.showNotification('Error:', result.error.message);
                 }
             }
         );
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    };

    //extract title from the document
    function extractTitle(str) {
        var flag = str.indexOf(">Title");
        if (flag != -1) {
            var begin = flag + 8;
            var end = str.indexOf("</", flag);
            return str.substring(begin, end);
        }
        else return '';
    };

    //extract a chapter from the HTML code
    function extractChapter(str, startChapter, stopChapter) {
        var flag = str.indexOf(startChapter);
        if (flag != -1 && str.indexOf(stopChapter) != -1) {
            var begin = str.indexOf("</h", flag) + 5;
            var flag2 = str.indexOf(stopChapter);
            var end = str.lastIndexOf('<h', flag2);
            return str.substring(begin, end);
        }
        else if (flag === -1) return '';
        else {
            var begin = str.indexOf("</h", flag) + 5;
            var end = str.indexOf('<h', begin);
            if (end < begin) end = str.length - 1;
            return str.substring(begin, end);
        }
    };

    //extract responsibilities and returns a list 
    function extractResponsibilities(str) {
        var respons = [3];
        // in some lay-outs they use 3 white spaces
        var test1 = str.indexOf('>Product Producer');
        var test2 = str.indexOf('>Product Producer');

        if (str.indexOf('>Product Producer') != -1) {
            var flag = str.indexOf(">Product Producer");
            var flag4 = 0; var flag2 = 0; var flag3 = 0;
            if (flag != -1) {
                var flagvalue = str.charAt(flag);
                var flag2 = str.indexOf("</td>", flag) + 10;
                var flag2value = str.charAt(flag2);
                var flag3 = str.indexOf(">", flag2) + 1;
                var flag3value = str.charAt(flag3);
                flag4 = str.indexOf("</td>", flag3);
                var flag4value = str.charAt(flag4);
                var producer = str.substring(flag3, flag4);
                respons[0] = producer;
            }
            else respons[0] = '';

            flag = str.indexOf(">Product Reviewer(s)", flag4);
            if (flag != -1) {
                flag2 = str.indexOf("</td>", flag) + 10;
                flag3 = str.indexOf(">", flag2) + 1;
                flag4 = str.indexOf("</td>", flag3);
                var reviewer = str.substring(flag3, flag4);
                respons[1] = reviewer;
            }
            else respons[1] = '';

            flag = str.indexOf(">Product Approver(s)", flag4);
            if (flag != -1) {
                flag2 = str.indexOf("</td>", flag) + 10;
                flag3 = str.indexOf(">", flag2) + 1;
                flag4 = str.indexOf("</td>", flag3);
                var approver = str.substring(flag3, flag4);
                respons[2] = approver;
            }
            else respons[2] = '';
            //app.showNotification(reviewer);
            return respons;
        }
        else {
            respons[0] = '';
            respons[1] = '';
            respons[2] = '';
            return respons
        }
    };

    //extract Development Skills and returns a matrix
    function extractDevSkills(str) {
        //app.showNotification('initialized devskills')
        if (str.indexOf('>Quality Skills Required') != -1) {
            var start = str.indexOf('>Quality Skills Required');
            var flag1 = str.indexOf('>Quality Skills Required');
            var flag2 = str.indexOf('>Quality Skills Required');
            var criteria = 0;

            while (str.indexOf('<tr>', flag1) < str.indexOf('</table>', start)) {
                criteria++;
                flag1 = str.indexOf('<tr>', flag1) + 3;
            }
            var devSkills = new Array(criteria);
            for (var j = 0; j < criteria; j++) {
                devSkills[j] = new Array(4);
                for (var i = 0; i < 4; i++) {
                    devSkills[j][i] = '';
                }
            }

            for (var i = 0; i < criteria; i++) {
                var flag3 = str.indexOf("<td", flag2);
                var flag4 = str.indexOf(">", flag3) + 1;
                var flag5 = str.indexOf("</td>", flag4);
                devSkills[i][0] = str.substring(flag4, flag5);

                flag3 = str.indexOf("<td", flag5);
                flag4 = str.indexOf(">", flag3) + 1;
                flag5 = str.indexOf("</td>", flag4);
                devSkills[i][1] = str.substring(flag4, flag5);


                flag3 = str.indexOf("<td", flag5);
                flag4 = str.indexOf(">", flag3) + 1;
                flag5 = str.indexOf("</td>", flag4);
                devSkills[i][2] = str.substring(flag4, flag5);


                flag3 = str.indexOf("<td", flag5);
                flag4 = str.indexOf(">", flag3) + 1;
                flag5 = str.indexOf("</td>", flag4);
                devSkills[i][3] = str.substring(flag4, flag5);

                flag2 = flag5;
            }
            return devSkills;
        }
        else return null
    };

    //test for completion of request
    function testForSuccess(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            app.showNotification('Error', asyncResult.error.message);
        }
    };

    //add text to Document
    function onAddContentHellowWorld() {
        var text = "Hello World";
        Office.context.document.setSelectedDataAsync(text, testForSuccess);
    };

    function adaptQualityCriteria(str, devSkills) {
        str.QualityCriteria.length = 0;
        var length = devSkills.length;
        for (var i = 0; i < length; i++) {
            var criteria = {
                "QualityCriteriaId": 12,
                "Criteria": devSkills[i][0],
                "Tolerance": devSkills[i][1],
                "Method": devSkills[i][2],
                "Skills": devSkills[i][3]
            }
            str.QualityCriteria.push(criteria);
        }

    };

    //convert html to text and add to document
    function onAddContentJson() {
        //var str = getProductDescription();
        // create HTML element
        var layout = "<head><style>p.MsoFooter, li.MsoFooter, div.MsoFooter{ margin:0in; margin-bottom:.0001pt; mso-pagination:widow-orphan; tab-stops:center 3.0in right 6.0in; font-size:12.0pt;} table {border-collapse: collapse;}table, th, td {border: 1px solid black;text-align: left; font-family: 'Calibri', 'sans-serif'} p,ol,ul{ font-family: 'Calibri', 'sans-serif'}</style></head>";
        var header = "<p class=MsoTitle>PRINCE2™- Product Description blabla</p>"
        //title
        var titledummy = "<h1>Title: ";
        var title = titledummy.concat(str.Title, "</h1>");
        //purpose
        var purpose = "<h1>Purpose</h1>";
        var purposeJson = str.Purpose;
        //composition
        var composition = "<h1>Composition</h1>";
        var compositionJson = str.Composition;
        //derivation
        var derivation = "<h1>Derivation</h1>";
        var derivationJson = str.Derivation;
        //format
        var format = "<h1>Format and Presentation</h1>";
        var formatJson = str.FormatPresentation;
        //skills
        var skills = "<h1>Development Skills Required</h1>";
        var div = $("<div>")
                .append(layout, header, title, purpose, purposeJson, composition, compositionJson, derivation,
                derivationJson, format, formatJson, skills);
        //tableQuality
        var tableSkills = "<table style='width:100%' id='table5' > <tr> <th>Quality Criteria </th> <th>Quality Tolerance </th> <th>Quality Method </th> <th>Quality Skills Required </th> </tr>";
        //checks how many criterias there are, write them out in table form
        for (var i = 0; i < str.QualityCriteria.length; i++) {
            tableSkills = tableSkills.concat("<tr><td>".concat(str.QualityCriteria[i].Criteria,
                "</td><td>", str.QualityCriteria[i].Tolerance, "</td><td>", str.QualityCriteria[i].Method,
                "</td><td>", str.QualityCriteria[i].Skills).concat("</td></tr>"));

        }
        tableSkills = tableSkills.concat("<tr><td></td><td></td><td></td><td></td></tr>")
        div.append(tableSkills.concat("</table>"));

        //Qualityresponsibility
        var responsibilities = "<h1>Quality Responsibilities</h1>"
        var tableResponsibility = "<table style='width:100%' id='table5'> <tr><th>Role</th><th>Responsible Individuals</th></tr><tr><th>Product Producer</th><td><p>"
            .concat(str.QualityResponsibility.Producer, "</p></td></tr><tr><th>Product Reviewer(s)</th><td>",
        str.QualityResponsibility.Reviewers,
        "</td></tr><tr><th>Product Approver(s)</th><td>",
        str.QualityResponsibility.Approvers, "</td></tr>");
        tableResponsibility = tableResponsibility.concat("</table>");
        div.append(responsibilities, tableResponsibility);

        // insert HTML into Word document
        Office.context.document.setSelectedDataAsync(div.html(), { coercionType: "html" }, testForSuccess);
    };

    //get HTML from a Word Document and add to document as text         
    function onGetHtml() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Html,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    Office.context.document.setSelectedDataAsync(result.value, testForSuccess);
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    };
})()