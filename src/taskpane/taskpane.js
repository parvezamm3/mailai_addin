/* global Office console */

export async function insertText(text) {
  // Write text to the cursor point in the compose surface.
  try {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
        console.log("Assync Result", asyncResult);
      }
    );
  } catch (error) {
    console.log("Error: " + error);
  }
}

// Office.onReady((info) => {
//     if (info.host === Office.HostType.Outlook) {
//         const resultString = sessionStorage.getItem('validationResult');
//         console.log(resultString)
//         if (resultString) {
//             const validationResult = JSON.parse(resultString);
//             // Display the message in your React component based on the result
//             if (validationResult.status === "fatal") {
//                 // Show a fatal error message in your UI
//             } else if (validationResult.status === "warning") {
//                 // Show a warning and a "Send Anyway" button
//             } else {
//                 // Show a success message
//                 console.log("success");
//             }
//             // Clear the stored data after using it
//             sessionStorage.removeItem('validationResult');
//         }
//     }
// });
