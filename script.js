Office.initialize = function () { };

try {
    Office.actions.associate("main", main);
} catch (e) {
    console.log("Unable to call 'Office.actions.associate'", e);
};

const PROMPT_URL = "https://www.google.com/";
let clickEvent;
let dialog;

// entry point
function main(event) {
    clickEvent = event;
    openDialog();
    return;
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
