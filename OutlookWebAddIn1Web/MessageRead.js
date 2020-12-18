'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
        });
    });

    var trimmed = []

    function requestFunction(j) {
        $.ajax({
            type: 'POST',
            url: 'https://192.168.1.52:8080/v2/phishbot',
            dataType: 'json',
            contentType: "application/json;",
            data: ('{"url": "' + trimmed[j] + '" }'),
            success: function (data) {
                console.log(data)
                var textchange = '#' + j + '.rank'
                var boxchange = '#' + j + '.mark'
                var progresschange = '#' + j + '.progress'
                var blacklistchange = '#' + j + '.blacklistMessage'
                var errorchange = '#' + j + '.errorMessage'
                $("#please").html(data);
                if (data == "BL") {
                    $(progresschange).css("display", "none")
                    $(blacklistchange).css("display", "block")
                    $(boxchange).css("display", "none")
                }

                //$(textchange).html(data)
                console.log(boxchange)
                var rank = Math.round(parseInt(data))
                $(textchange).html(rank.toString())
                rank = (rank * 2) + 6
                $(boxchange).css("left", "" + rank.toString() + "px")
                // var json = $.parseJSON(data);
                // alert(json.html);
            },
            error: function (data) {
                console.log('error')
                var textchange = '#' + j + '.rank'
                var boxchange = '#' + j + '.mark'
                var progresschange = '#' + j + '.progress'
                var blacklistchange = '#' + j + '.blacklistMessage'
                var errorchange = '#' + j + '.errorMessage'
                $(progresschange).css("display", "none")
                $(boxchange).css("display", "none")
                $(errorchange).css("display", "block")
                //alert(json.error);
            }

        });
    }

    function loadItemProps(item) {
        // Write message property values to the task pane
        //$(.'x_WordSection1')
        console.log("HALLO")
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        Office.context.mailbox.item.body.getAsync(
            "html",
            { asyncContext: "This is passed to the callback" },
            function callback(result) {
                // Do something with the result.
                var regxpattern = /(href="http[s]?:\/\/)[(www\.)?a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)(")/gi
                //var regxpattern = /(href=")(http|ftp|https):\/ \/ ([\w +?\.\w +]) + ([a - zA - Z0 - 9\~\!\@\#\$\%\^\&\*\(\)_\-\=\+\\\/\?\.\:\;\'\,]*)?(")/g
                var raw = result.value
                // var content = raw.outerHTML
                var results = raw.match(regxpattern)
                //var links = raw.link
                console.log(raw)

                console.table(results)
                console.log("hello")
                //var table = document.getElementById('tt')

                
                for (var i = 0; i < results.length; i++) {
                    trimmed[i] = results[i].substring(6, results[i].length - 1)
                    //$('#tt tr:last').after("<tr> <td>" + trimmed[i] + "</td><th><div style = 'background: linear-gradient(to right, #ff9966 0%, #ff99cc 100%);'></div></th><th>Quantitative risk</th></tr>")
                    $('.entry').last().after("<div class='entry' id='" + i + "'><p class='url'>" + trimmed[i] + "</p><div class='progress' id = '" + i + "'></div><p class='blacklistMessage' id='" + i + "'>This URL has been Blacklisted by your company due to its high risk.</p><p class='errorMessage' id='" + i + "'>Error: PhishBot was unable to retrieve this address.</p><div class='mark' id='" + i + "'><div class='rank' id='" + i + "'>0.0</div><i class='fa fa-caret-up'></i></div></div>")
                    //$('.entry').last().append("<div class='entry'><p>" + trimmed[i] + "</p><div class='progress'></div><p>Quantitative ranking</p></div>")
                }
                console.table(trimmed)
                /*
                $.ajax({
                    type: 'POST',
                    url: 'https://192.168.1.52:8080/v2/phishbot',
                    dataType: 'json',
                    contentType: "application/json;",
                    data: ('{"url": "https://forms.office.com/Pages/ResponsePage.aspx?i…KtvK91_Py9XlUNjFXVVlYNUYwOU1LVVVLVzJaU0NUNEE0VS4u" }'),
                    success: function (data) {
                        console.log(data)
                        
                        var textchange = '#' + j + '.rank'
                        var boxchange = '#' + j + '.mark'
                        var progresschange = '#' + j + '.progress'
                        var blacklistchange = '#' + j + '.blacklistMessage'
                        $("#please").html(data);
                        if (data == "BL") {
                            $(progresschange).css("display", "none")
                            $(blacklistchange).css("display", "block")
                        }

                        $(textchange).html(data)
                        console.log(boxchange)
                        var rank = Math.round(parseInt(data))
                        rank = (rank * 2) + 6
                        $(boxchange).css("left", "" + rank.toString() + "px")
                        
                        // var json = $.parseJSON(data);
                        // alert(json.html);
                    },
                    error: function (data) {
                        console.log('error')
                        //alert(json.error);
                    }

                }); */
                
                for (var j = 0; j < trimmed.length; j++) {
                    requestFunction(j)
                }
                
            });
        //$('#item-id').text(item.itemId);

    }
})();