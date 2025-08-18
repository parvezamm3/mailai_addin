const FLASK_BASE_URL = "https://equipped-externally-stud.ngrok-free.app";
function onMessageSend(event) {
  // Prevent the email from being sent immediately
//   event.completed({ allowEvent: false }); 
  console.log('onMessageSend activated')
  // You can now call a function from your App.jsx to handle the validation
  // For simplicity, we'll assume a global function is exposed, but a better
  // pattern is to use a shared service or event bus.
  handleOutgoingEmail(event);
// Office.context.mailbox.item.body.getAsync(
//     "text",
//     { asyncContext: event },
//     getBodyCallback
//   );

}

function getBodyCallback(asyncResult){
  const event = asyncResult.asyncContext;
  console.log(event);
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
    console.log("body");
    console.log(body);
  } else {
    const message = "Failed to get body text";
    console.error(message);
    console.log(asyncResult);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  const matches = hasMatches(body);
  if (matches) {
    Office.context.mailbox.item.getAttachmentsAsync(
      { asyncContext: event },
      getAttachmentsCallback);
  } else {
    event.completed({ allowEvent: true });
  }
}

function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }
  console.log("Has Matches");

  const arrayOfTerms = ["send", "picture", "document", "attachment"];
  for (let index = 0; index < arrayOfTerms.length; index++) {
    const term = arrayOfTerms[index].trim();
    const regex = RegExp(term, 'i');
    if (regex.test(body)) {
      return true;
    }
  }
  console.log("HasMatches finished");

  return false;
}

function getAttachmentsCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (let i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({ allowEvent: true });
        return;
      }
    }

    event.completed({
      allowEvent: false,
      errorMessage: "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown: "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook)."
    });
  } else {
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like you're forgetting to include an attachment.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown: "Looks like you're forgetting to include an attachment.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook)."
    });
  }
}

// function sendAnywayFunction(event) {
//   // Hide the warning notification
//   Office.context.mailbox.item.notificationMessages.removeAsync("warning");
  
//   // This will re-trigger the send event, but this time we let it go through.
//   // The event object needs to be completed with allowEvent: true.
//   // This is a simplified approach, in a real scenario you might need to
//   // store the original event and complete it here.
//   event.completed({ allowEvent: true });
// }

async function handleOutgoingEmail (event){
    // Display a loading message in the Outlook UI
    console.log("validating mail");
    // console.log(event);
    Office.context.mailbox.item.notificationMessages.addAsync("progress", {
      type: "informationalMessage",
      message: "Checking your email for potential issues...",
      icon: "Icon.16x16",
      persistent: false,
    });
    
    try {
      const item = Office.context.mailbox.item;
    //   console.log(item);
      const [
            subjectResult, 
            bodyResult, 
            recipientsResult,
            senderResult, 
            attachmentsResult 
        ] = await Promise.all([
            new Promise(resolve => item.subject.getAsync(result => resolve(result))),
            new Promise(resolve => item.body.getAsync(Office.CoercionType.Text, result => resolve(result))),
            new Promise(resolve => item.to.getAsync(result => resolve(result))),
            new Promise(resolve => item.from.getAsync(result => resolve(result))),
            new Promise(resolve => item.getAttachmentsAsync(result => resolve(result))),
        ]);
        const subject = subjectResult.value;
        const body = bodyResult.value;
        const recipients = recipientsResult.value.map(rec => rec.emailAddress);
        const sender = senderResult.value.emailAddress;
        const conversationId = item.conversationId;
        const attachments = attachmentsResult.value;
    //   console.log("Subject:", subject);
    //     console.log("Sender:", sender);
    //     console.log("Conversation ID:", conversationId);
    //   console.log("Recipients", recipients);
    //   console.log("Body",body);
    // console.log(attachments[0].name);
    // console.log(Office.context.mailbox.userProfile);
      const outgoingPayload = {
        sende:sender,
        subject: subject,
        body: body,
        recipients: recipients,
        conv_id:conversationId,
        attachments:attachments
      };

      // Call the new backend endpoint for outgoing email validation
      const response = await fetch(`${FLASK_BASE_URL}/validate_outgoing`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(outgoingPayload),
      });
    //   console.log("Response")
    //   console.log(response);
      const result = await response.json();
      Office.context.mailbox.item.notificationMessages.removeAsync("progress"); // Remove loading message

      if (!response.ok) {
        // Backend returned an error, assume fatal issue
        Office.context.mailbox.item.notificationMessages.addAsync("error", {
          type: "errorMessage",
          message: `Backend validation failed: ${result.message}`,
          icon: "Icon.32x32",
          persistent: true,
        });
        event.completed({ allowEvent: false });
        return;
      }
      
      // Handle the validation result from the backend
      if (result.status === "fatal") {
        // A fatal issue was found, block the email and show the user the error
        Office.context.mailbox.item.notificationMessages.addAsync("fatalIssue", {
          type: "errorMessage",
          message: `🚫 Cannot send: ${result.message}`,
          icon: "Icon.32x32",
          persistent: true,
        });
        event.completed({ allowEvent: false });
      } else if (result.status === "warning") {
        // A warning was found, show a notification with an option to send anyway
        Office.context.mailbox.item.notificationMessages.addAsync("warning", {
          type: "informationalMessage",
          message: `⚠️ Warning: ${result.message}`,
          icon: "Icon.16x16",
          persistent: true,
          actions: [
            {
              id: "sendAnyway",
              label: "Send Anyway",
              actionType: "executeFunction",
              functionName: "sendAnywayFunction"
            }
          ]
        });
        
        // This is a temporary placeholder. You'll need to define sendAnywayFunction
        // in your `commands.js` to complete the send event.
        event.completed({ allowEvent: false }); 
      } else {
        // No issues found, allow the email to be sent
        event.completed({ allowEvent: true });
    //     event.completed({
    //   allowEvent: false,
    //   errorMessage: "Testuing.",
    //   // TIP: In addition to the formatted message, it's recommended to also set a
    //   // plain text message in the errorMessage property for compatibility on
    //   // older versions of Outlook clients.
    //   errorMessageMarkdown: "Testing."
    // });
      }
    } catch (error) {
      console.error("Error during outgoing email validation:", error);
      Office.context.mailbox.item.notificationMessages.removeAsync("progress");
      Office.context.mailbox.item.notificationMessages.addAsync("error", {
        type: "errorMessage",
        message: `An unexpected error occurred during validation: ${error.message}`,
        icon: "Icon.32x32",
        persistent: true,
      });
      event.completed({ allowEvent: true }); // Fail-safe: allow send on error
    }
  };

// Register the function with Outlook
Office.actions.associate("onMessageSend", onMessageSend);
// Office.actions.associate("sendAnywayFunction", sendAnywayFunction);