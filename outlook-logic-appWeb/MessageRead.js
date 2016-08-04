/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var bug_url = "<logic-app-url>";
    var pbi_url = "<logic-app-url>";
    var task_url = "<logic-app-url>";
    var wunderlist_url = "<logic-app-url>";

    var messageBanner;
    var spinner;
    var title;

    var url_array = [bug_url, task_url, pbi_url, "na", wunderlist_url];
    var fabricComponent = {};

    function reqListener() {
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner.showBanner();
        var data = JSON.parse(this.responseText);
        $('#spinner').hide();
        $('#link-to-item').attr('href', data['url']);
        $('#link-to-item').text(data['display']);
        $('#input-field-link').val(data['url']);
        $('#result-link').show();
        $('#input-field-link').focus();
        $('#input-field-link').select();
    }

    function send_request(item, result, url) {
        $('#main-grid').hide();
        $('#spinner').show();
        var xhr = new XMLHttpRequest();
        xhr.open("POST", url);
        xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        xhr.addEventListener('load', reqListener);
        item.bodyHTML = result.value;
        item.form_priority = $('#priority').val();
        item.form_title = $('#title').val();
        title = item.form_title;
        item.form_assigned = $('#assigned').val();
        xhr.send(JSON.stringify(item));
    }

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            spinner = fabricComponent.Spinner(document.getElementById('spinner'));
            var item = Office.context.mailbox.item;
            
            $('#spinner').hide();
            $('#result-link').hide();
            $('#copy-message').hide();
            $('#title').val(item.subject);

            //loadProps();
            $('#logicapp-button').click(function ()
            {
                Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
                    send_request(Office.context.mailbox.item, result, url_array[document.getElementById('type').selectedIndex]);
                });
            });

            document.getElementById("copy-button").addEventListener("click", function () {
                copyToClipboard(document.getElementById('input-field-link'));
                $('#copy-message').show();
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

    /**
     * @param {HTMLDOMElement} target - The element the Spinner will attach itself to.
     */

    

    fabricComponent.Spinner = function (target) {

        var _target = target;
        var eightSize = 0.2;
        var circleObjects = [];
        var animationSpeed = 90;
        var interval;
        var spinner;
        var numCircles;
        var offsetSize;
        var fadeIncrement = 0;
        var parentSize = 20;

        /**
         * @function start - starts or restarts the animation sequence
         * @memberOf fabric.Spinner
         */
        function start() {
            stop();
            interval = setInterval(function () {
                var i = circleObjects.length;
                while (i--) {
                    _fade(circleObjects[i]);
                }
            }, animationSpeed);
        }

        /**
         * @function stop - stops the animation sequence
         * @memberOf fabric.Spinner
         */
        function stop() {
            clearInterval(interval);
        }

        //private methods

        function _init() {
            _setTargetElement();
            _setPropertiesForSize();
            _createCirclesAndArrange();
            _initializeOpacities();
            start();
        }

        function _initializeOpacities() {
            var i = 0;
            var j = 1;
            var opacity;
            fadeIncrement = 1 / numCircles;

            for (i; i < numCircles; i++) {
                var circleObject = circleObjects[i];
                opacity = (fadeIncrement * j++);
                _setOpacity(circleObject.element, opacity);
            }
        }

        function _fade(circleObject) {
            var opacity = _getOpacity(circleObject.element) - fadeIncrement;

            if (opacity <= 0) {
                opacity = 1;
            }

            _setOpacity(circleObject.element, opacity);
        }

        function _getOpacity(element) {
            return parseFloat(window.getComputedStyle(element).getPropertyValue("opacity"));
        }

        function _setOpacity(element, opacity) {
            element.style.opacity = opacity;
        }

        function _createCircle() {
            var circle = document.createElement('div');
            circle.className = "ms-Spinner-circle";
            circle.style.width = circle.style.height = parentSize * offsetSize + "px";
            return circle;
        }

        function _createCirclesAndArrange() {

            var angle = 0;
            var offset = parentSize * offsetSize;
            var step = (2 * Math.PI) / numCircles;
            var i = numCircles;
            var circleObject;
            var radius = (parentSize - offset) * 0.5;

            while (i--) {
                var circle = _createCircle();
                var x = Math.round(parentSize * 0.5 + radius * Math.cos(angle) - circle.clientWidth * 0.5) - offset * 0.5;
                var y = Math.round(parentSize * 0.5 + radius * Math.sin(angle) - circle.clientHeight * 0.5) - offset * 0.5;
                spinner.appendChild(circle);
                circle.style.left = x + 'px';
                circle.style.top = y + 'px';
                angle += step;
                circleObject = { element: circle, j: i };
                circleObjects.push(circleObject);
            }
        }

        function _setPropertiesForSize() {
            if (spinner.className.indexOf("large") > -1) {
                parentSize = 28;
                eightSize = 0.179;
            }

            offsetSize = eightSize;
            numCircles = 8;
        }

        function _setTargetElement() {
            //for backwards compatibility
            if (_target.className.indexOf("ms-Spinner") === -1) {
                spinner = document.createElement("div");
                spinner.className = "ms-Spinner";
                _target.appendChild(spinner);
            } else {
                spinner = _target;
            }
        }

        _init();

        return {
            start: start,
            stop: stop
        };
    };

    

})();

