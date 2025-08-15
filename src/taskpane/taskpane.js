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
