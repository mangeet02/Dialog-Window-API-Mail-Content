Office.initialize = function () { };

try {
    Office.actions.associate("main", main);
} catch (e) {
    console.log("Unable to call 'Office.actions.associate'", e);
};

const PROMPT_URL = "https://mangeet02.github.io/Dialog-Window-API-Mail-Content/form.html";
let clickEvent;
let dialog;

// entry point
function main(event) {
    clickEvent = event;
    openDialog();
    loadItemProps(Office.context.mailbox.item);
    return;
}

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
                console.error('Error -> No mail selected?', result.error.message);
            }
        }
    );
}

function openDialog() {
    Office.context.ui.displayDialogAsync(PROMPT_URL, { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

function dialogCallback(asyncResult) {
    dialog = asyncResult.value;

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
}

function eventHandler(arg) {
    console.log(`Received event from prompt dialog.`);
    console.table(arg);
    clickEvent.completed();
}

/*
async function messageHandler(arg) {
    console.log(`Received message "${arg.message}" from prompt dialog.`);
    dialog.close();
    switch (arg.message) {
        case "cancel":
            dialog.close();
            clickEvent.completed();
            break;
        case "move":
            let itemId = Office.context.mailbox.item.itemId;
            let folderId = "JunkEmail";
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, 
                async (result) => {
                    const token = result.value;
                    let res = await moveEmail(itemId, folderId, token);
                    console.log(`API response from request to move item to ${folderId}:`);
                    console.table(res);
                    clickEvent.completed();
                });
            break;
        default:
            break;
    }
}

async function moveEmail(itemId, folderId, token) {
    console.log(`Moving email with item ID ${itemId} to ${folderId}.`);
    const itemUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/";
    const restItemId =  Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    const junkUrl = itemUrl + restItemId + "/move";

    let res = await fetch(junkUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json;charset=utf-8',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
            "DestinationId": folderId
        })
      });
    return res;
} 

*/
