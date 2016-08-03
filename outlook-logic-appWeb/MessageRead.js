/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var bug_url = "https://prod-04.westus.logic.azure.com:443/workflows/3bc2ac3aacfd4792abd2d199793210e3/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YPmLoGzv_qgzcevEZeoWWMBBf1UHEMqLA5RSkOQVHmM";
    var pbi_url = "https://prod-12.westus.logic.azure.com:443/workflows/112f6e7ef61b433097247290bb8704ee/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_7eaLwcQ-yurddUynwQpy2hInI94RtZeSJox0ciBxvE";
    var task_url = "https://prod-09.westus.logic.azure.com:443/workflows/9d6f38b4b8bc4193a408f5350f30da89/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=C_KihSSK0rSkgzdyN4rvrIOSn1VYIcHiRjvXKN3yx3A";

    var messageBanner;

    function reqListener() {
        $('#message').text('Sent to Logic App');
    }

    function send_request(item, result, url) {
        var xhr = new XMLHttpRequest();
        xhr.open("POST", url);
        xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        xhr.addEventListener('load', reqListener);
        item.bodyHTML = result.value;
        xhr.send(JSON.stringify(item));
    }

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            //loadProps();
            $('#bug').click(function ()
            {
                Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
                    send_request(Office.context.mailbox.item, result, bug_url);
                });
            });

            $('#task').click(function () {
                Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
                    send_request(Office.context.mailbox.item, result, task_url);
                });
            });

            $('#pbi').click(function () {
                Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
                    send_request(Office.context.mailbox.item, result, pbi_url);
                });
            });

        });
    }



})();