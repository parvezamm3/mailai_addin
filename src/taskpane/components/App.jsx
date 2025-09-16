import React, { useState, useEffect } from "react";
import { Spinner, SpinnerSize, Stack, Text, MessageBar, MessageBarType } from "@fluentui/react";
import ToggleSwitch from "./ToggleSwitch";
import SuggestedReplies from "./SuggestedReplies";
import { initializeIcons } from "@fluentui/react/lib/Icons";

// Initialize Fluent UI icons at the top of your app
initializeIcons();

const FLASK_BASE_URL = "https://equipped-externally-stud.ngrok-free.app";

const App = () => {
  const [isLoading, setIsLoading] = useState(true);
  const [isGenerateImportance, setGenerateImportance] = useState(false);
  const [isGenerateReplies, setGenerateReplies] = useState(false);
  const [isSpam, setIsSpam] = useState(false);
  const [isAnalysisNeeded, setIsAnalysisNeeded] = useState(true);
  const [isPromotional, setIsPromotional] = useState(false);
  const [statusMessage, setStatusMessage] = useState();
  const [importanceScore, setImportanceScore] = useState();
  const [importanceDescription, setImportanceDescription] = useState();
  const [suggestedReplies, setSuggestedReplies] = useState([]);
  const [categories, setCategories] = useState([]);
  const [summary, setSummary] = useState("");
  const [errorMessage, setErrorMessage] = useState("");
  const [isAuthorized, setIsAuthorized] = useState(true);
  const [ownerEmail, setOwnerEmail] = useState("");

  useEffect(() => {
    Office.onReady((info) => {
      // Log all properties of the info object for detailed debugging
      if (info.host === Office.HostType.Outlook) {
        // console.log("Office host is Outlook. Loading email details...");
        loadEmailDetails();
        try {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            loadEmailDetails,
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // console.log("Successfully added ItemChanged event handler.");
              } else {
                console.error(
                  `Failed to add ItemChanged event handler: ${asyncResult.error.message}`
                );
              }
            }
          );
        } catch (error) {
          console.error(`Error trying to add ItemChanged handler: ${error.message}`);
        }
      } else {
        console.log("Add-in not running in Outlook. Displaying fallback message.");
        // setAnalysisResult("This add-on is designed for Outlook. Please open it within Outlook.");
        setIsLoading(false);
      }
    });
    // --- NEW: Cleanup function for useEffect ---
    // Remove the event handler when the component unmounts to prevent memory leaks.
    return () => {
      if (Office.context && Office.context.mailbox) {
        try {
          Office.context.mailbox.removeHandlerAsync(
            Office.EventType.ItemChanged,
            { handler: loadEmailDetails },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully removed ItemChanged event handler.");
              } else {
                console.error(
                  `Failed to remove ItemChanged event handler: ${asyncResult.error.message}`
                );
              }
            }
          );
        } catch (error) {
          console.error(`Error trying to remove ItemChanged handler: ${error.message}`);
        }
      }
    };
  }, []);

  const loadEmailDetails = async () => {
    console.log("loadEmailDetails function started.");
    setIsLoading(true); // Ensure loading state is active
    setIsAuthorized(true); // Assume authorized until proven otherwise
    try {
      const item = Office.context.mailbox.item;
      // console.log("Office.context.mailbox.item:", item);

      if (!item) {
        console.log("No mail item found. Setting analysis result and ending loading.");
        setErrorMessage("このアドオンの機能を使用するには、メールを開いてください。");
        setIsLoading(false);
        return;
      }
      const emailSubject = item.subject;
      const emailSender = item.sender.emailAddress;
      const messageId = item.itemId;
      const convId = item.conversationId;
      const userEmail = Office.context.mailbox.userProfile.emailAddress;
      let tempOwnerEmail = userEmail;
      try {
        Office.context.mailbox.item.getSharedPropertiesAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            // The method succeeded, so the item is in a shared mailbox or shared folder.
            // console.log("Getting result");
            const sharedProperties = asyncResult.value;
            tempOwnerEmail = sharedProperties.owner;
            // console.log("After async");
            console.log(userEmail, tempOwnerEmail);
          }
        });
      } catch (error) {
        // console.log(error);
      }
      // console.log(userEmail, tempOwnerEmail);
      item.body.getAsync(Office.CoercionType.Text, async (asyncResult) => {
        console.log("item.body.getAsync result status:", asyncResult.status);
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = asyncResult.value;
          // console.log(userEmail, tempOwnerEmail);
          // console.log("Inside body");
          setOwnerEmail(tempOwnerEmail);
          // --- Fetch analysis and preferences from Flask backend ---
          const payloadAnalysis = {
            user_id: userEmail,
            ownerEmail: tempOwnerEmail,
            provider: "outlook",
            sender: emailSender,
            subject: emailSubject,
            body: emailBody,
            conv_id: convId,
            message_id: messageId,
          };
          try {
            const responseAnalysis = await fetch(`${FLASK_BASE_URL}/dashboard_data`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(payloadAnalysis),
            });
            // console.log(responseAnalysis);
            const dataAnalysis = await responseAnalysis.json();
            setStatusMessage(dataAnalysis.status_message);
            // console.log(dataAnalysis.status_message);
            // if (responseAnalysis.status === 202) {

            //   console.log(dataAnalysis);
            // }
            if (responseAnalysis.status === 401) {
              const errorText = await responseAnalysis.text();
              console.warn(`User not authorized. Displaying authorization prompt.: ${errorText}`);
              setIsAuthorized(false); // Set authorization to false
              setErrorMessage("AI 機能にアクセスするには承認が必要です。");
            } else if (responseAnalysis.status === 200) {
              // const dataAnalysis = await responseAnalysis.json();
              setErrorMessage(dataAnalysis.message);
              setIsAnalysisNeeded(dataAnalysis.is_analysis_needed);
              setIsSpam(dataAnalysis.is_spam);
              // console.log(isSpam);
              setIsPromotional(dataAnalysis.is_promotional);
              setImportanceScore(dataAnalysis.importance_score);
              setImportanceDescription(dataAnalysis.importance_description);

              setSuggestedReplies(dataAnalysis.replies);
              console.log(Object.keys(dataAnalysis.replies).length);
              setCategories(dataAnalysis.categories);
              setSummary(dataAnalysis.summary);
              setGenerateImportance(
                dataAnalysis.preferences?.enable_importance_generation || false
              );

              setGenerateReplies(dataAnalysis.preferences?.enable_reply_generation || false);
            }
            // else {
            //   const errorText = responseAnalysis.message;
            //   setErrorMessage(errorText);
            //   console.log("Error fetching dashboard data:", errorText);
            // }
          } catch (error) {
            setErrorMessage(`分析のためにバックエンドに接続できませんでした: ${error.message}`);
            console.error("Network error during analysis fetch:", error);
          }

          setIsLoading(false);
          console.log("isLoading set to false. UI should now fully render.");
        } else {
          // Handle error if email body cannot be retrieved
          setErrorMessage(`メール本文の取得中にエラーが発生しました: ${asyncResult.error.message}`);
          console.error("Error in item.body.getAsync:", asyncResult.error);
          setIsLoading(false);
        }
      });
    } catch (error) {
      console.error("General error in loadEmailDetails:", error);
      // setAnalysisResult(`予期しないエラーが発生しました: ${error.message}`);
      setIsLoading(false);
      setIsAuthorized(false);
    }
  };

  const handleToggleChange = async (fieldName, isChecked) => {
    // console.log(`Toggle change detected: ${fieldName} to ${isChecked}`);
    if (fieldName === "enable_importance_generation") {
      setGenerateImportance(isChecked);
    } else if (fieldName === "enable_reply_generation") {
      setGenerateReplies(isChecked);
    }

    const payload = {
      user_id: Office.context.mailbox.userProfile.emailAddress,
      platform: "outlook",
      enable_importance_generation: isGenerateImportance,
      enable_reply_generation: isGenerateReplies,
    };

    if (fieldName === "enable_importance_generation") {
      payload.enable_importance_generation = isChecked;
    } else if (fieldName === "enable_reply_generation") {
      payload.enable_reply_generation = isChecked;
    }

    // console.log("Saving preferences to Flask:", payload);
    try {
      const response = await fetch(`${FLASK_BASE_URL}/save_preferences`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        if (fieldName === "enable_importance_generation") setGenerateImportance(!isChecked);
        if (fieldName === "enable_reply_generation") setGenerateReplies(!isChecked);
        // const errorText = await response.text();
      } else {
        console.log("Preferences successfully saved to Flask.");
      }
    } catch (error) {
      if (fieldName === "enable_importance_generation") setGenerateImportance(!isChecked);
      if (fieldName === "enable_reply_generation") setGenerateReplies(!isChecked);
      console.error("Network error saving preferences:", error);
    }
  };

  const handleRefresh = () => {
    console.log("Refresh button clicked. Reloading email details...");
    loadEmailDetails(); // Re-run the main loading function
  };

  const handleReplyClick = (replyText) => {
    // console.log("Attempting to display reply form with text:", replyText.substring(0, 50) + "...");
    const replyHtml = replyText.replace(/\n/g, "<br />");
    const replyOptions = {
      htmlBody: replyHtml,
    };
    Office.context.mailbox.item.displayReplyForm(replyOptions);
    // Office.context.mailbox.item.displayReplyForm(replyText);
  };

  async function toggleAnalysisNeeded(isChecked, ownerEmail) {
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    const convId = item.conversationId;
    const payload = {
      user_email: ownerEmail,
      conv_id: convId,
      message_id: messageId,
      is_analysis_needed: isChecked,
      platform: "outlook",
    };

    try {
      const response = await fetch(`${FLASK_BASE_URL}/set_analysis_needed`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        handleRefresh();
        console.log(payload);
      }
    } catch (error) {
      console.error("Network error during toggleAnalysisNeeded method:", error);
      // Add a user-friendly notification if needed
    }
  }

  async function toggleIsPromotional(isChecked, ownerEmail) {
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    const convId = item.conversationId;
    const payload = {
      user_email: ownerEmail,
      conv_id: convId,
      message_id: messageId,
      is_promotional: isChecked,
      platform: "outlook",
    };

    try {
      const response = await fetch(`${FLASK_BASE_URL}/set_is_promotional`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        handleRefresh();
      }
    } catch (error) {
      console.error("Network error during toggleIsPromotional method:", error);
      // Add a user-friendly notification if needed
    }
  }

  async function toggleIsSpam(isChecked, ownerEmail) {
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    const convId = item.conversationId;
    const payload = {
      user_email: ownerEmail,
      conv_id: convId,
      message_id: messageId,
      is_spam: isChecked,
      platform: "outlook",
    };
    // console.log(payload);

    try {
      const response = await fetch(`${FLASK_BASE_URL}/set_is_spam`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        handleRefresh();
      }
    } catch (error) {
      console.error("Network error during toggleIsSpam method:", error);
      // Add a user-friendly notification if needed
    }
  }

  if (isLoading) {
    return (
      <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20, textAlign: "center" } }}>
        <Spinner label="Loading..." size={SpinnerSize.large} />
      </Stack>
    );
  }

  const FormattedText = ({ text }) => {
    return (
      <>
        {text.split("\n").map((line, index) => (
          <React.Fragment key={index}>
            {line}
            {index < text.split("\n").length - 1 && <br />}
          </React.Fragment>
        ))}
      </>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20 } }}>
      {/* <Text variant="xxLarge" styles={{ root: { fontWeight: "bold" } }}>
        MailAI
      </Text> */}
      <Text variant="medium">メールを分析して返信の提案を得る</Text>
      {/* Refresh Button */}
      <button
        onClick={handleRefresh}
        className="w-full bg-gray-200 text-gray-800 py-2 px-4 rounded-md hover:bg-gray-300 focus:outline-none focus:ring-2 focus:ring-gray-400 focus:ring-opacity-50 mb-4"
      >
        コンテンツをリフレッシュ
      </button>
      {errorMessage && (
        <>
          <Stack>
            <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>
          </Stack>
        </>
      )}

      {/* Conditional Rendering based on Authorization */}
      {isAuthorized ? (
        <>
          {/* Feature Control Section */}
          <Stack horizontal tokens={{ childrenGap: 5 }}>
            {/* First Column */}
            <Stack
              tokens={{ childrenGap: 5, padding: 5 }}
              styles={{ root: { border: "1px solid #ccc", borderRadius: 8 } }}
            >
              <Text variant="large" styles={{ root: { fontWeight: "bold" } }}>
                機能
              </Text>
              <ToggleSwitch
                label="重要度分析"
                checked={isGenerateImportance}
                onToggle={(checked) => handleToggleChange("enable_importance_generation", checked)}
              />
              <ToggleSwitch
                label="返信生成"
                checked={isGenerateReplies}
                onToggle={(checked) => handleToggleChange("enable_reply_generation", checked)}
              />
            </Stack>

            {/* Second Column (Empty example, but you can add your content here) */}
            <Stack
              tokens={{ childrenGap: 5, padding: 5 }}
              styles={{ root: { border: "1px solid #ccc", borderRadius: 8 } }}
            >
              <Text variant="medium" styles={{ root: { fontWeight: "bold" } }}>
                フィードバック
              </Text>
              <ToggleSwitch
                label="分析必要"
                checked={isAnalysisNeeded}
                onToggle={(checked) => toggleAnalysisNeeded(checked, ownerEmail)}
              />
              <ToggleSwitch
                label="スパム"
                checked={isSpam}
                onToggle={(checked) => toggleIsSpam(checked, ownerEmail)}
              />
              <ToggleSwitch
                label="プロモーション"
                checked={isPromotional}
                onToggle={(checked) => toggleIsPromotional(checked, ownerEmail)}
              />
              {/* Add more content, like another set of toggles or controls */}
            </Stack>
          </Stack>

          {/* Analysis Result Section */}
          <Text variant="large" styles={{ root: { fontWeight: "bold" } }}>
            分析
          </Text>
          {statusMessage && statusMessage != "分析は正常に完了しました。" && (
            <>
              <Stack>
                <MessageBar messageBarType={MessageBarType.info}>{statusMessage}</MessageBar>
              </Stack>
            </>
          )}
          {!isSpam && (
            <>
              {categories && categories.length > 0 && (
                <Stack>
                  <p>
                    <b>カテゴリー</b> : {categories.join(", ")}
                  </p>
                  <p>
                    <b>要約 : </b>
                    {summary}{" "}
                  </p>
                </Stack>
              )}

              {isGenerateImportance && importanceScore && (
                <Stack
                  tokens={{ childrenGap: 10, padding: 15 }}
                  styles={{ root: { border: "1px solid #ccccccff", borderRadius: 8 } }}
                >
                  <p>
                    <b>重要度スコア</b> : {importanceScore}
                  </p>
                  <p>
                    <b>説明 : </b>
                    {importanceDescription}
                  </p>
                  {/* <FormattedText text={analysisResult} /> */}
                </Stack>
              )}

              {isGenerateReplies && Object.keys(suggestedReplies).length > 0 && (
                // <>

                <SuggestedReplies replies={suggestedReplies} onReplyClick={handleReplyClick} />
                // </>
              )}
            </>
          )}
        </>
      ) : (
        // Authorization Section
        <Stack
          tokens={{ childrenGap: 10, padding: 15 }}
          styles={{ root: { border: "1px solid #ccc", borderRadius: 8 } }}
        >
          <Text variant="large" styles={{ root: { fontWeight: "bold" } }}>
            Authorization Required
          </Text>
          <Text variant="medium">
            Please authorize your Microsoft 365 account to use the AI Assistant features.
          </Text>
          <button
            onClick={() => window.open(`${FLASK_BASE_URL}/outlook-authorize`, "_blank")}
            className="w-full bg-blue-500 text-white py-2 px-4 rounded-md hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50"
          >
            Authorize with Microsoft
          </button>
        </Stack>
      )}
    </Stack>
  );
};

export default App;
