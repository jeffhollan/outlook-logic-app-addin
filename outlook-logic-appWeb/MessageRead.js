/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var bug_url = "https://prod-04.westus.logic.azure.com:443/workflows/3bc2ac3aacfd4792abd2d199793210e3/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YPmLoGzv_qgzcevEZeoWWMBBf1UHEMqLA5RSkOQVHmM";
    var pbi_url = "https://prod-12.westus.logic.azure.com:443/workflows/112f6e7ef61b433097247290bb8704ee/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_7eaLwcQ-yurddUynwQpy2hInI94RtZeSJox0ciBxvE";
    var task_url = "https://prod-09.westus.logic.azure.com:443/workflows/9d6f38b4b8bc4193a408f5350f30da89/triggers/manual/run?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=C_KihSSK0rSkgzdyN4rvrIOSn1VYIcHiRjvXKN3yx3A";

    var messageBanner;


    var url_array = [bug_url, task_url, pbi_url, "na"];


    function reqListener() {
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner = new fabric.MessageBanner(element);
        messageBanner.showBanner();
        $('#main-grid').hide();
        $('#result-link').show();
        $('#input-field-link').focus();
        $('#input-field-link').select();
    }

    function send_request(item, result, url) {
        var xhr = new XMLHttpRequest();
        xhr.open("POST", url);
        xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        xhr.addEventListener('load', reqListener);
        item.bodyHTML = result.value;
        item.form_priority = $('#priority').val();
        item.form_title = $('#title').val();
        item.form_assigned = $('#assigned').val();
        xhr.send(JSON.stringify(item));
    }

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            //document.getElementById("copyButton").addEventListener("click", function () {
            //    copyToClipboard(document.getElementById('input-field-link'));
            //    $('#link-div').focus();
            //});

            var item = Office.context.mailbox.item;
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            $('#result-link').hide()

            $('#title').val(item.subject);

            //loadProps();
            $('#logicapp-button').click(function ()
            {
                Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
                    send_request(Office.context.mailbox.item, result, url_array[document.getElementById('type').selectedIndex]);
                });
            });

        });
    }

    function copyToClipboard(elem) {
        // create hidden text element, if it doesn't already exist
        var targetId = "_hiddenCopyText_";
        var isInput = elem.tagName === "INPUT" || elem.tagName === "TEXTAREA";
        var origSelectionStart, origSelectionEnd;
        if (isInput) {
            // can just use the original source element for the selection and copy
            target = elem;
            origSelectionStart = elem.selectionStart;
            origSelectionEnd = elem.selectionEnd;
        } else {
            // must use a temporary form element for the selection and copy
            target = document.getElementById(targetId);
            if (!target) {
                var target = document.createElement("textarea");
                target.style.position = "absolute";
                target.style.left = "-9999px";
                target.style.top = "0";
                target.id = targetId;
                document.body.appendChild(target);
            }
            target.textContent = elem.textContent;
        }
        // select the content
        var currentFocus = document.activeElement;
        target.focus();
        target.setSelectionRange(0, target.value.length);

        // copy the selection
        var succeed;
        try {
            succeed = document.execCommand("copy");
        } catch (e) {
            succeed = false;
        }
        // restore original focus
        if (currentFocus && typeof currentFocus.focus === "function") {
            currentFocus.focus();
        }

        if (isInput) {
            // restore prior selection
            elem.setSelectionRange(origSelectionStart, origSelectionEnd);
        } else {
            // clear temporary content
            target.textContent = "";
        }
        return succeed;
    }


})();

