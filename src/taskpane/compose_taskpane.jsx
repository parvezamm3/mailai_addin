import React, { useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import { 
    Stack,
    Text,
    PrimaryButton
    
} from "@fluentui/react";
import {
    initializeIcons
} from "@fluentui/react/lib/Icons";

initializeIcons();

const customButtonStyles = {
  root: {
    backgroundColor: '#784086ff',
    border: 'none',
    borderRadius: '999px',
    padding: '10px 15px', 
  },
  rootHovered: {
    backgroundColor: "#690505ff",
  },
  label: {
    color: 'white',
    fontWeight: 'bold',
  },
};

const ComposeTaskPane = () => {
//   const [message, setMessage] = useState(null);
    const [errorMessages, setErrorMessages] = new useState({})
  useEffect(() => {
    // This hook runs when the component is mounted
    Office.onReady(() => {
      const resultString = sessionStorage.getItem('validationResult');
    //   console.log(resultString);
     if (resultString ) {
    //    const validationResult = JSON.parse(resultString);
    const validationResult = JSON.parse(resultString);
        let error_messages = {}
        if (validationResult['sensitive_data']['has_sensitive_data']){
            error_messages["機密データが含まれています"]=validationResult['sensitive_data']['comment']
        }
        if (validationResult['attachments']['has_missing_attachments']){
            error_messages["添付ファイルが不足しています"]=validationResult['attachments']['comment']
        }
        if (validationResult['grammatical_errors']['has_errors']){
            error_messages["文法上の誤りがあります"]=validationResult['grammatical_errors']['comment']
        }
        if (validationResult['spelling_mistakes']['has_mistakes']){
            error_messages["スペルミスがあります"]=validationResult['spelling_mistakes']['comment']
        }
        if (validationResult['best_practices']['is_not_followed']){
            error_messages["ベストプラクティスに準拠していません"]=validationResult['best_practices']['comment']
        }
        setErrorMessages(error_messages);
       // Clean up sessionStorage after reading
       sessionStorage.removeItem('validationResult');
     }
    });
  }, []);

  const handleSendAnyway = () => {
    // 1. Set the flag in sessionStorage.
    sessionStorage.setItem('sendAnyway', 'true');
    
    // 2. Remove the notification message.
    Office.context.mailbox.item.notificationMessages.removeAsync("warning");
    
    // 3. Close the task pane. This resumes the OnMessageSend event,
    // which will now see the flag and allow the send.
    Office.context.ui.closeContainer();
};

  return (
    <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20 }}>

    {Object.entries(errorMessages).map(([key, value]) => {
        return (
          <Stack key={key}>
            <Text variant="medium" styles={{ root: { fontWeight: "bold", color:"red" } }}>
              {key}
            </Text>
            <Text variant="medium">{value}</Text>
          </Stack>
          
        );
      }
    )}
    <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 20 } }}>
                <PrimaryButton styles={customButtonStyles} onClick={handleSendAnyway} text="問題なさそうです！確認せずに送信を続行！" />
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