Office.initialize = function () {
}

var dialog;
var clickEvent;

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function justTest(event) {

  //window.location.href = "https://www.google.co.th";
  window.open("https://www.pttgcgroup.com/");
  event.completed();

  //openDialog();
  //  doSomethingAndShowDialog(event);

  
}

function openDialog() {
    Office.context.ui.displayDialogAsync("~remoteAppUrl/MessageRead.html",
        { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Domain is not trusted");
                break;
            case 12005:
                showNotification("HTTPS is required");
                break;
            case 12007:
                showNotification("A dialog is already opened.");
                break;
            default:
                showNotification(asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;
        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        //dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        //dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}

function showNotification(text) {
    //writeToDoc(text);
    //Required, call event.completed to let the platform know you are done processing.
    clickEvent.completed();
}

function doSomethingAndShowDialog(event) {
    clickEvent = event;
    //writeToDoc("Ribbon button clicked.");
    openDialog();
}

