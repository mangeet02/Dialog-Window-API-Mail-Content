'use strict';

(function () {

    Office.onReady(function () {
        $(document).ready(function () {
            loadItemProps(Office.context.mailbox.item);
        });
    });

    function loadItemProps(item) {

        $('#item-subject').text(item.subject);

        Office.context.mailbox.item.body.getAsync(
            "text",
            { asyncContext: "This is passed to the callback" },
            function callback(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var emailBody = result.value;
                    document.getElementById("body").textContent = emailBody;
                } else {
                    console.error('Errore nel recupero del corpo dell\'email:', result.error.message);
                }
            }
        );

    }

})();
