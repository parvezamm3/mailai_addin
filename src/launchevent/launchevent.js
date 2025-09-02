const FLASK_BASE_URL = "https://equipped-externally-stud.ngrok-free.app";
function onMessageSend(event) {
  // sessionStorage.setItem('validationResult', JSON.stringify({ status: "loading", message: "Checking your email for potential issues..." }));
  console.log("Get Item : ",sessionStorage.getItem("sendAnyway"))
  if (sessionStorage.getItem("sendAnyway") === "true") {
    sessionStorage.removeItem("sendAnyway"); // Clean up the flag
    console.log("Cleaned Item : ",sessionStorage.getItem("sendAnyway"))
    event.completed({ allowEvent: true }); // Allow the send to proceed
    return;
  }
  handleOutgoingEmail(event);
}

async function handleOutgoingEmail(event) {
  // Display a loading message in the Outlook UI
//   console.log("validating mail");
  // console.log(event);
  Office.context.mailbox.item.notificationMessages.addAsync("progress", {
    type: "informationalMessage",
    message: "潜在的な問題がないかメールを確認しています...",
    icon: "Icon.16x16",
    persistent: false,
  });
  // const TASK_PANE_URL = "https://localhost:3000/compose_taskpane.html";

  try {
    const item = Office.context.mailbox.item;
      // console.log(item);
    const [
      subjectResult,
      bodyResult,
      recipientsResult,
      ccResult,
      bccResult,
      senderResult,
      attachmentsResult,
    ] = await Promise.all([
      new Promise((resolve) => item.subject.getAsync((result) => resolve(result))),
      new Promise((resolve) =>
        item.body.getAsync(Office.CoercionType.Text, { bodyMode: Office.MailboxEnums.BodyMode.HostConfig }  , (result) => resolve(result))
      ),
      new Promise((resolve) => item.to.getAsync((result) => resolve(result))),
      new Promise((resolve) => item.cc.getAsync((result) => resolve(result))),
      new Promise((resolve) => item.bcc.getAsync((result) => resolve(result))),
      new Promise((resolve) => item.from.getAsync((result) => resolve(result))),
      new Promise((resolve) => item.getAttachmentsAsync((result) => resolve(result))),
    ]);
    const subject = subjectResult.value;
    const body = bodyResult.value;
    const recipients = recipientsResult.value.map((rec) => rec.emailAddress);
    const cc = ccResult.value.map((rec) => rec.emailAddress);
    const bcc = bccResult.value.map((rec) => rec.emailAddress);
    const sender = senderResult.value.emailAddress;
    const conversationId = item.conversationId;
    const attachments = attachmentsResult.value;
    // console.log(body);
    // console.log(attachments);
    // console.log(Office.context.mailbox.userProfile.emailAddress);
    const outgoingPayload = {
      // message_id : item.id,
      sender: sender,
      subject: subject,
      body: body,
      recipients: recipients,
      cc: cc,
      bcc: bcc,
      conv_id: conversationId,
      attachments: attachments,
      email_address:Office.context.mailbox.userProfile.emailAddress
    };

    // Call the new backend endpoint for outgoing email validation
    const response = await fetch(`${FLASK_BASE_URL}/validate_outgoing`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(outgoingPayload),
    });
    const result = await response.json();
    const response_data_string = result.data;
    //     // console.log(response_data_string);
    const response_data = JSON.parse(response_data_string);
    // console.log(response_data)
    // const response_data = {};
    if (response_data && Object.keys(response_data).length != 0) {
      Office.context.mailbox.item.notificationMessages.removeAsync("progress");
      let has_missing_attachments = response_data["attachments"]["has_missing_attachments"];
      let is_best_practices_followed = response_data["best_practices"]["is_not_followed"];
      let has_grmmatical_errors = response_data["grammatical_errors"]["has_errors"];
      let has_sensitive_data = response_data["sensitive_data"]["has_sensitive_data"];
      let has_spelling_mistakes = response_data["spelling_mistakes"]["has_mistakes"];

      if (
        has_missing_attachments ||
        is_best_practices_followed ||
        has_grmmatical_errors ||
        has_sensitive_data ||
        has_spelling_mistakes
      ) {
        sessionStorage.setItem("validationResult", JSON.stringify(response_data));
        event.completed({
          allowEvent: false,
          errorMessage: "メール本文に誤りがあるようです。メールを送信する前に修正してください。",
          cancelLabel: "詳細",
          commandId: "composeOpenPaneButton",
        });
      }
      event.completed({
        allowEvent: true,
      });
    }
    event.completed({
      allowEvent: false,
      errorMessage: "データ処理中にエラーが発生しました。",
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
  } catch (error) {
    console.error("Error during outgoing email validation:", error);
    // On unexpected errors, allow the email to be sent as a fail-safe.
    Office.context.mailbox.item.notificationMessages.removeAsync("progress");
    event.completed({ 
        allowEvent: false, 
        errorMessage: "送信メールの検証中にエラーが発生しました。", 
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser});
  }
}

// Register the function with Outlook
Office.actions.associate("onMessageSend", onMessageSend);
// Office.actions.associate("sendAnywayFunction", sendAnywayFunction);
