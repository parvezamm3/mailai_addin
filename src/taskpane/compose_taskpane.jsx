import React, { useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import { Stack, Text, PrimaryButton } from "@fluentui/react";
import { initializeIcons } from "@fluentui/react/lib/Icons";

initializeIcons();

const customButtonStyles = {
  root: {
    backgroundColor: "#784086ff",
    border: "none",
    borderRadius: "999px",
    padding: "10px 15px",
  },
  rootHovered: {
    backgroundColor: "#690505ff",
  },
  label: {
    color: "white",
    fontWeight: "bold",
  },
};

const ComposeTaskPane = () => {
  //   const [message, setMessage] = useState(null);
  const [errorMessages, setErrorMessages] = useState({});
  const [correctedVersion, setCorrectedVersion] = useState({});
  useEffect(() => {
    // This hook runs when the component is mounted
    Office.onReady(() => {
      const resultString = sessionStorage.getItem("validationResult");
      if (resultString) {
        const validationResult = JSON.parse(resultString);
        let tempErrorMessages = {};

        if (validationResult.sensitive_data.has_sensitive_data) {
          tempErrorMessages["機密データが含まれています"] = validationResult.sensitive_data.comment;
        }
        if (validationResult.attachments.has_missing_attachments) {
          tempErrorMessages["添付ファイルが不足しています"] = validationResult.attachments.comment;
        }
        if (validationResult.grammatical_errors.has_errors) {
          tempErrorMessages["文法上の誤りがあります"] = validationResult.grammatical_errors.comment;
        }
        if (validationResult.spelling_mistakes.has_mistakes) {
          tempErrorMessages["スペルミスがあります"] = validationResult.spelling_mistakes.comment;
        }
        if (validationResult.best_practices.is_not_followed) {
          tempErrorMessages["ベストプラクティスに準拠していません"] =
            validationResult.best_practices.comment;
        }
        setErrorMessages(tempErrorMessages);

        if (validationResult.corrected_version) {
          setCorrectedVersion({
            subject: validationResult.corrected_version.subject,
            body: validationResult.corrected_version.body.replace(/\n/g, "<br />"),
          });
        }
        // Clean up sessionStorage after reading
        sessionStorage.removeItem("validationResult");
      }
    });
  }, []);

  const handleSendAnyway = () => {
    // 1. Set the flag in sessionStorage.
    sessionStorage.setItem("sendAnyway", "true");

    // 2. Remove the notification message.
    Office.context.mailbox.item.notificationMessages.removeAsync("warning");

    // 3. Close the task pane. This resumes the OnMessageSend event,
    // which will now see the flag and allow the send.
    Office.context.ui.closeContainer();
  };

  const handleReplyClick = () => {
    // console.log("Attempting to display reply form with text:", replyText.substring(0, 50) + "...");
    // const replyHtml = correctedVersion.body.replace(/\n/g, "<br />");
    console.log("Sending mail");
    Office.context.mailbox.item.subject.setAsync(correctedVersion.subject, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set subject: " + asyncResult.error.message);
      }
    });
    // console.log("Subject ok");

    Office.context.mailbox.item.body.setAsync(
      correctedVersion.body,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set body: " + asyncResult.error.message);
        }
      }
    );
    // console.log("Body ok");

    // You might also want to close the task pane after updating
    sessionStorage.setItem("sendAnyway", "true");
    Office.context.ui.closeContainer();
  };

  return (
    <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20 }}>
      {Object.entries(errorMessages).map(([key, value]) => {
        return (
          <Stack key={key}>
            <Text variant="medium" styles={{ root: { fontWeight: "bold", color: "red" } }}>
              {key}
            </Text>
            <Text variant="medium">{value}</Text>
          </Stack>
        );
      })}
      <Stack>
        <Text variant="medium" styles={{ root: { fontWeight: "bold", color: "black" } }}>
          修正された返信
        </Text>
        <Text variant="medium" styles={{ root: { color: "black" } }}>
          件名 : {correctedVersion.subject}
        </Text>
        <Text variant="medium" styles={{ root: { color: "black" } }}>
          ボディー : {correctedVersion.body}
        </Text>
        <PrimaryButton
          text="この内容で返信する"
          onClick={handleReplyClick} // Call the corrected handler
          styles={customButtonStyles}
        />
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 20 } }}>
        <PrimaryButton
          styles={customButtonStyles}
          onClick={handleSendAnyway}
          text="問題なさそうです！そのまま送信！"
        />
        {/* <DefaultButton onClick={handleCancelSend} text="Cancel" /> */}
      </Stack>
    </Stack>
  );
};

// This is the file that the webpack config will target.
// It renders the ComposeTaskPane component into the DOM.
ReactDOM.createRoot(document.getElementById("container")).render(
  <React.StrictMode>
    <ComposeTaskPane />
  </React.StrictMode>
);
