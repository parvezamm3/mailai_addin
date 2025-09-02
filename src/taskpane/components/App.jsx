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
  const [analysisResult, setAnalysisResult] = useState("Loading analysis and preferences...");
  const [isGenerateImportance, setGenerateImportance] = useState(false);
  const [isGenerateSummaryAndCategorize, setGenerateSummaryAndCategorize] = useState(false);
  const [isGenerateReplies, setGenerateReplies] = useState(false);
  const [is_spam, set_is_spam] = useState(false);
  const [is_malicious, set_is_malicious] = useState(false);
  const [suggestedReplies, setSuggestedReplies] = useState([]);
  const [category, setCategory] = useState("");
  const [summary, setSummary] = useState("");
  // const [emailDetails, setEmailDetails] = useState(null);
  const [errorMessage, setErrorMessage] = useState("");
  const [isAuthorized, setIsAuthorized] = useState(true);

  // const [outgoingValidation, setOutgoingValidation] = useState(null); // NEW state for outgoing validation

  useEffect(() => {
    // Expose the function globally so the commands.js handler can call it

    // console.log("App component mounted. Calling Office.onReady...");
    Office.onReady((info) => {
      // Log all properties of the info object for detailed debugging
      // console.log("Office.onReady fired. info object:", info);

      // Changed condition: Only check for host type. If onReady fires, it's generally initialized enough.
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
        setAnalysisResult("This add-on is designed for Outlook. Please open it within Outlook.");
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
    // console.log("loadEmailDetails function started.");
    setIsLoading(true); // Ensure loading state is active
    setIsAuthorized(true); // Assume authorized until proven otherwise
    try {
      const item = Office.context.mailbox.item;
      // console.log("Office.context.mailbox.item:", item);

      if (!item) {
        console.log("No mail item found. Setting analysis result and ending loading.");
        setErrorMessage("このアドオンの機能を使用するには、メールを開いてください。");
        setIsLoading(false);
        // setIsAuthorized(false);
        return;
      }
      // for
      // console.log(item.to);
      const emailSubject = item.subject;
      const emailSender = item.sender.emailAddress;
      const messageId = item.itemId;
      const convId = item.conversationId;
      const userEmail = Office.context.mailbox.userProfile.emailAddress;
      let ownerEmail = userEmail;
      try {
        Office.context.mailbox.item.getSharedPropertiesAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            // The method succeeded, so the item is in a shared mailbox or shared folder.
            const sharedProperties = asyncResult.value;
            ownerEmail = sharedProperties.owner;
          }
        });
      } catch (error) {
        // console.log(error);
      }
      item.body.getAsync(Office.CoercionType.Text, async (asyncResult) => {
        console.log("item.body.getAsync result status:", asyncResult.status);
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = asyncResult.value;

          // --- Fetch analysis and preferences from Flask backend ---
          const payloadAnalysis = {
            user_id: userEmail,
            ownerEmail: ownerEmail,
            provider: "outlook",
            sender: emailSender,
            subject: emailSubject,
            body: emailBody,
            conv_id: convId,
            message_id: messageId,
          };
          // console.log("Sending analysis request to Flask:", payloadAnalysis);
          try {
            const responseAnalysis = await fetch(`${FLASK_BASE_URL}/dashboard_data`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(payloadAnalysis),
            });
            // console.log(responseAnalysis);

            // console.log("Flask analysis response status:", responseAnalysis.status);
            if (responseAnalysis.status === 401) {
              const errorText = await responseAnalysis.text();
              console.warn(`User not authorized. Displaying authorization prompt.: ${errorText}`);
              setIsAuthorized(false); // Set authorization to false
              setErrorMessage("AI 機能にアクセスするには承認が必要です。");
            } else if (responseAnalysis.ok) {
              const dataAnalysis = await responseAnalysis.json();
              // console.log("Flask analysis data received:", dataAnalysis);
              setErrorMessage(dataAnalysis.message);
              // console.log(dataAnalysis.message);
              set_is_spam(dataAnalysis.is_spam);
              set_is_malicious(dataAnalysis.is_malicious);
              setAnalysisResult(dataAnalysis.analysis_result || "結果が返されませんでした。");
              // console.log(dataAnalysis.replies);
              setSuggestedReplies(dataAnalysis.replies);
              setCategory(dataAnalysis.category);
              setSummary(dataAnalysis.summary);
              setGenerateImportance(
                dataAnalysis.preferences?.enable_importance_generation || false
              );
              setGenerateSummaryAndCategorize(
                dataAnalysis.preferences?.enable_summarization_and_categorization || false
              );
              setGenerateReplies(dataAnalysis.preferences?.enable_reply_generation || false);
            } else {
              const errorText = responseAnalysis.message;
              setErrorMessage(errorText);
              console.log("Error fetching dashboard data from Flask:", errorText);
            }
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
      setAnalysisResult(`予期しないエラーが発生しました: ${error.message}`);
      setIsLoading(false);
      setIsAuthorized(false);
    }
  };

  // const generateAnalysis = async () => {
  //   try {
  //     const item = Office.context.mailbox.item;
  //     const messageId = item.itemId;
  //     const convId = item.conversationId;
  //     const userEmail = Office.context.mailbox.userProfile.emailAddress;
  //     let ownerEmail = userEmail;
  //     try {
  //       Office.context.mailbox.item.getSharedPropertiesAsync((asyncResult) => {
  //       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //         // The method succeeded, so the item is in a shared mailbox or shared folder.
  //         const sharedProperties = asyncResult.value;
  //         ownerEmail = sharedProperties.owner;
  //       }
  //     });
  //     } catch (error) {
  //       console.log(error);
  //     }
  //     item.body.getAsync(Office.CoercionType.Text, async (asyncResult) => {
  //       // console.log("item.body.getAsync result status:", asyncResult.status);
  //       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //         const emailBody = asyncResult.value;

  //         // --- Fetch analysis and preferences from Flask backend ---
  //         const payloadAnalysis = {
  //           ownerEmail:ownerEmail,
  //           message_id: messageId,
  //         };
  //         // console.log("Sending analysis request to Flask:", payloadAnalysis);
  //         try {
  //           const responseAnalysis = await fetch(`${FLASK_BASE_URL}/generate_analysis_outlook`, {
  //             method: "POST",
  //             headers: { "Content-Type": "application/json" },
  //             body: JSON.stringify(payloadAnalysis),
  //           });

  //           if (responseAnalysis.ok) {
  //             const dataAnalysis = await responseAnalysis.json();
  //             // console.log("Flask analysis data received:", dataAnalysis);
  //             set_is_spam(dataAnalysis.is_spam);
  //             set_is_malicious(dataAnalysis.is_malicious);
  //             setAnalysisResult(
  //               dataAnalysis.analysis_result || "結果が返されませんでした"
  //             );
  //           } else {
  //             const errorText = await responseAnalysis.text();
  //             setAnalysisResult(`Error from backend (${responseAnalysis.status}): ${errorText}`);
  //             console.error("Error fetching dashboard data from Flask:", errorText);
  //           }
  //         } catch (error) {
  //           setAnalysisResult(`Failed to connect to backend for analysis: ${error.message}`);
  //           console.error("Network error during analysis fetch:", error);
  //         }

  //         // Only fetch replies if authorized
  //         if (isAuthorized) {
  //           // Check isAuthorized *after* the dashboard_data fetch
  //           // console.log("Fetching suggested replies...");
  //           const replies = await getSuggestedReplies(
  //             item,
  //             ownerEmail
  //           );
  //           setSuggestedReplies(replies);
  //           // console.log("Suggested replies state updated. Count:", replies.length);
  //         } else {
  //           setSuggestedReplies([]); // Clear replies if not authorized
  //         }

  //         setIsLoading(false);
  //         console.log("isLoading set to false. UI should now fully render.");
  //       } else {
  //         // Handle error if email body cannot be retrieved
  //         setAnalysisResult(`Error retrieving email body: ${asyncResult.error.message}`);
  //         console.error("Error in item.body.getAsync:", asyncResult.error);
  //         setIsLoading(false);
  //       }
  //     });
  //   } catch (error) {
  //     console.error("General error in loadEmailDetails:", error);
  //     setAnalysisResult(`An unexpected error occurred: ${error.message}`);
  //     setIsLoading(false);
  //     setIsAuthorized(false);
  //   }
  // }

  const handleToggleChange = async (fieldName, isChecked) => {
    // console.log(`Toggle change detected: ${fieldName} to ${isChecked}`);
    if (fieldName === "enable_importance_generation") {
      setGenerateImportance(isChecked);
    } else if (fieldName === "enable_summarization_and_categorization") {
      setGenerateSummaryAndCategorize(isChecked);
    } else if (fieldName === "enable_reply_generation") {
      setGenerateReplies(isChecked);
    }

    const payload = {
      user_id: Office.context.mailbox.userProfile.emailAddress,
      platform: "outlook",
      enable_importance_generation: isGenerateImportance,
      enable_summarization_and_categorization: isGenerateSummaryAndCategorize,
      enable_reply_generation: isGenerateReplies,
    };

    if (fieldName === "enable_importance_generation") {
      payload.enable_importance_generation = isChecked;
    } else if (fieldName === "enable_summarization_and_categorization") {
      payload.enable_summarization_and_categorization = isChecked;
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
        if (fieldName === "enable_summarization_and_categorization")
          setGenerateSummaryAndCategorize(!isChecked);
        if (fieldName === "enable_reply_generation") setGenerateReplies(!isChecked);
        const errorText = await response.text();
        // console.error("Failed to save preferences to Flask:", errorText);
        // Add a user-friendly notification if needed
      } else {
        console.log("Preferences successfully saved to Flask.");
      }
    } catch (error) {
      if (fieldName === "enable_importance_generation") setGenerateImportance(!isChecked);
      if (fieldName === "enable_summarization_and_categorization")
        setGenerateSummaryAndCategorize(!isChecked);
      if (fieldName === "enable_reply_generation") setGenerateReplies(!isChecked);
      console.error("Network error saving preferences:", error);
      // Add a user-friendly notification if needed
    }
  };

  // const getSuggestedReplies = async (item, userEmail) => {
  //   const SUGGEST_REPLY_URL = `${FLASK_BASE_URL}/suggest_reply`;
  //   let fetchedReplies = [];
  //   // console.log("Initiating getSuggestedReplies...");

  //   try {

  //     const payload = {
  //       user_id: userEmail,
  //       conv_id: item.conversationId,
  //       message_id: item.itemId,
  //     };
  //     // console.log("Sending reply suggestion request to Flask:", payload);

  //     const response = await fetch(SUGGEST_REPLY_URL, {
  //       method: "POST",
  //       headers: { "Content-Type": "application/json" },
  //       body: JSON.stringify(payload),
  //     });

  //     // console.log("Flask reply suggestion response status:", response.status);
  //     if (response.ok) {
  //       const data = await response.json();
  //       // console.log("Flask reply suggestion data received:", data);
  //       if (data.suggested_replies && Array.isArray(data.suggested_replies)) {
  //         fetchedReplies = data.suggested_replies;
  //         setCategory(data.category);
  //       } else {
  //         console.warn("Backend did not return an array of suggested replies in expected format.");
  //       }
  //     } else {
  //       const errorText = await response.text();
  //       console.error(
  //         `Error fetching suggested replies from backend (${response.status}): ${errorText}`
  //       );
  //     }
  //   } catch (error) {
  //     console.error("Network error fetching suggested replies:", error);
  //   }
  //   // console.log("Finished getSuggestedReplies. Found:", fetchedReplies.length, "replies.");
  //   return fetchedReplies;
  // };

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

  async function setNotMalicious() {
    const item = Office.context.mailbox.item;
    const messageId = item.itemId;
    const convId = item.conversationId;
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    const payload = {
      user_email: userEmail,
      conv_id: convId,
      message_id: messageId,
      platform: "outlook",
    };

    try {
      const response = await fetch(`${FLASK_BASE_URL}/not_malicious`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        handleRefresh();
      }
    } catch (error) {
      console.error("Network error during setNotMalicious method:", error);
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
      <Text variant="xxLarge" styles={{ root: { fontWeight: "bold" } }}>
        MailAI
      </Text>
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

          <Stack
            tokens={{ childrenGap: 10, padding: 15 }}
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
              label="要約を生成と分類"
              checked={isGenerateSummaryAndCategorize}
              onToggle={(checked) =>
                handleToggleChange("enable_summarization_and_categorization", checked)
              }
            />
            <ToggleSwitch
              label="返信生成"
              checked={isGenerateReplies}
              onToggle={(checked) => handleToggleChange("enable_reply_generation", checked)}
            />
          </Stack>

          {/* Analysis Result Section */}
          <Text variant="large" styles={{ root: { fontWeight: "bold" } }}>
            分析
          </Text>
          {!is_spam && !is_malicious ? (
            <>
              {isGenerateImportance && (
                <Stack
                  tokens={{ childrenGap: 10, padding: 15 }}
                  styles={{ root: { border: "1px solid #ccccccff", borderRadius: 8 } }}
                >
                  <FormattedText text={analysisResult} />
                </Stack>
              )}
              {isGenerateSummaryAndCategorize && (
                <>
                  <p>
                    <b>カテゴリー</b> : {category}
                  </p>
                  <p>
                    <b>要約 : </b>
                    {summary}{" "}
                  </p>
                </>
              )}

              {isGenerateReplies && (
                // <>

                <SuggestedReplies replies={suggestedReplies} onReplyClick={handleReplyClick} />
                // </>
              )}
            </>
          ) : (
            <>
              <MessageBar messageBarType={MessageBarType.error}>
                このメールはスパムまたは悪意のあるコンテンツとして検出されました。
              </MessageBar>
              <button onClick={setNotMalicious}>安全そうに見える</button>
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
