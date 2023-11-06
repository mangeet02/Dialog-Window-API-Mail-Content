Office.initialize = function () { };

try {
    Office.actions.associate("main", main);
} catch (e) {
    console.log("Unable to call 'Office.actions.associate'", e);
};

const PROMPT_URL = "https://fantasy.dunkest.com/#/signup";
let clickEvent;
let dialog;

//https://mangeet02.github.io/Dialog-Window-API-Mail-Content/form.html

function main(event) {
    clickEvent = event;
    openDialog();
    return;
}

function openDialog() {
    Office.context.ui.displayDialogAsync(PROMPT_URL, { height: 50, width: 50, displayInIframe: false }, dialogCallback);
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
