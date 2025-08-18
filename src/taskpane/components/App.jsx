import React, { useState, useEffect } from "react";
import { Panel, Pivot, PivotItem, Spinner, SpinnerSize, Stack, Text } from "@fluentui/react";
import ToggleSwitch from "./ToggleSwitch";
import SuggestedReplies from "./SuggestedReplies";
import { initializeIcons } from "@fluentui/react/lib/Icons";

// Initialize Fluent UI icons at the top of your app
initializeIcons();

const FLASK_BASE_URL = "https://equipped-externally-stud.ngrok-free.app";

const App = () => {
  const [isLoading, setIsLoading] = useState(true);
  const [analysisResult, setAnalysisResult] = useState("Loading analysis and preferences...");
  const [importanceEnabled, setImportanceEnabled] = useState(false);
  const [generationEnabled, setGenerationEnabled] = useState(false);
  const [suggestedReplies, setSuggestedReplies] = useState([]);
  const [emailDetails, setEmailDetails] = useState(null); 
  const [isAuthorized, setIsAuthorized] = useState(true);

  const [outgoingValidation, setOutgoingValidation] = useState(null); // NEW state for outgoing validation

  useEffect(() => {
    // Expose the function globally so the commands.js handler can call it

    console.log("App component mounted. Calling Office.onReady...");
    Office.onReady((info) => {
      // Log all properties of the info object for detailed debugging
      // console.log("Office.onReady fired. info object:", info);

      // Changed condition: Only check for host type. If onReady fires, it's generally initialized enough.
      if (info.host === Office.HostType.Outlook) {
        console.log("Office host is Outlook. Loading email details...");
        loadEmailDetails();
        try {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            loadEmailDetails,
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully added ItemChanged event handler.");
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
    console.log("loadEmailDetails function started.");
    setIsLoading(true); // Ensure loading state is active
    setIsAuthorized(true); // Assume authorized until proven otherwise
    try {
      const item = Office.context.mailbox.item;
      // console.log("Office.context.mailbox.item:", item);

      if (!item) {
        console.log("No mail item found. Setting analysis result and ending loading.");
        setAnalysisResult("Please open an email to use this add-on's functionality.");
        setIsLoading(false);
        setIsAuthorized(false);
        return;
      }

      const emailSubject = item.subject;
      const emailSender = item.sender.emailAddress;
      const messageId = item.itemId;
      const convId = item.conversationId;
      const userEmail = Office.context.mailbox.userProfile.emailAddress;

      // console.log("Email Subject:", emailSubject);
      // console.log("Email Sender:", emailSender);
      // console.log("Message ID:", messageId);
      // console.log("User Email:", userEmail);

      // Get the plain text body asynchronously
      item.body.getAsync(Office.CoercionType.Text, async (asyncResult) => {
        // console.log("item.body.getAsync result status:", asyncResult.status);
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = asyncResult.value;
          // console.log("Email Body successfully retrieved. Length:", emailBody.length);

          setEmailDetails({
            subject: emailSubject,
            body: emailBody,
            sender: emailSender,
            convId: convId,
            messageId: messageId,
            userEmail: userEmail,
          });

          // --- Fetch analysis and preferences from Flask backend ---
          const payloadAnalysis = {
            user_id: userEmail,
            platform: "outlook",
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

            // console.log("Flask analysis response status:", responseAnalysis.status);
            if (responseAnalysis.status === 401) {
              console.warn("User not authorized. Displaying authorization prompt.");
              setIsAuthorized(false); // Set authorization to false
              setAnalysisResult("Authorization required to access AI features.");
            } else if (responseAnalysis.ok) {
              const dataAnalysis = await responseAnalysis.json();
              // console.log("Flask analysis data received:", dataAnalysis);
              setAnalysisResult(
                dataAnalysis.analysis_result || "No specific analysis returned from backend."
              );
              setImportanceEnabled(dataAnalysis.preferences?.enable_importance || false);
              setGenerationEnabled(dataAnalysis.preferences?.enable_generation || false);
              // console.log(
              //   "Analysis and preferences state updated. Importance:",
              //   importanceEnabled,
              //   "Confidential:",
              //   generationEnabled
              // );
            } else {
              const errorText = await responseAnalysis.text();
              setAnalysisResult(`Error from backend (${responseAnalysis.status}): ${errorText}`);
              console.error("Error fetching dashboard data from Flask:", errorText);
            }
          } catch (error) {
            setAnalysisResult(`Failed to connect to backend for analysis: ${error.message}`);
            console.error("Network error during analysis fetch:", error);
          }

          // Only fetch replies if authorized
          if (isAuthorized) {
            // Check isAuthorized *after* the dashboard_data fetch
            console.log("Fetching suggested replies...");
            const replies = await getSuggestedReplies(
              item,
              userEmail,
              emailSubject,
              emailBody,
              emailSender
            );
            setSuggestedReplies(replies);
            // console.log("Suggested replies state updated. Count:", replies.length);
          } else {
            setSuggestedReplies([]); // Clear replies if not authorized
          }

          setIsLoading(false);
          console.log("isLoading set to false. UI should now fully render.");
        } else {
          // Handle error if email body cannot be retrieved
          setAnalysisResult(`Error retrieving email body: ${asyncResult.error.message}`);
          console.error("Error in item.body.getAsync:", asyncResult.error);
          setIsLoading(false);
        }
      });
    } catch (error) {
      console.error("General error in loadEmailDetails:", error);
      setAnalysisResult(`An unexpected error occurred: ${error.message}`);
      setIsLoading(false);
      setIsAuthorized(false);
    }
  };

  const handleToggleChange = async (fieldName, isChecked) => {
    // console.log(`Toggle change detected: ${fieldName} to ${isChecked}`);
    if (fieldName === "enable_importance") {
      setImportanceEnabled(isChecked);
    } else if (fieldName === "enable_generation") {
      setGenerationEnabled(isChecked);
    }

    const payload = {
      user_id: Office.context.mailbox.userProfile.emailAddress,
      platform: "outlook",
      enable_importance: importanceEnabled,
      enable_generation: generationEnabled,
    };

    if (fieldName === "enable_importance") {
      payload.enable_importance = isChecked;
    } else if (fieldName === "enable_generation") {
      payload.enable_generation = isChecked;
    }

    // console.log("Saving preferences to Flask:", payload);
    try {
      const response = await fetch(`${FLASK_BASE_URL}/save_preferences`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        if (fieldName === "enable_importance") setImportanceEnabled(!isChecked);
        if (fieldName === "enable_generation") setGenerationEnabled(!isChecked);
        const errorText = await response.text();
        // console.error("Failed to save preferences to Flask:", errorText);
        // Add a user-friendly notification if needed
      } else {
        console.log("Preferences successfully saved to Flask.");
      }
    } catch (error) {
      if (fieldName === "enable_importance") setImportanceEnabled(!isChecked);
      if (fieldName === "enable_generation") setGenerationEnabled(!isChecked);
      console.error("Network error saving preferences:", error);
      // Add a user-friendly notification if needed
    }
  };

  const getSuggestedReplies = async (item, userEmail, subject, body, sender) => {
    const SUGGEST_REPLY_URL = `${FLASK_BASE_URL}/suggest_reply`;
    let fetchedReplies = [];
    // console.log("Initiating getSuggestedReplies...");

    try {
      const emailBodyText = await new Promise((resolve) =>
        item.body.getAsync(Office.CoercionType.Text, (r) => resolve(r.value))
      );
      // console.log("Email body for reply suggestion:", emailBodyText.length, "characters");

      const payload = {
        user_id: userEmail,
        platform: "outlook",
        conv_id: item.conversationId,
        message_id: item.itemId,
        subject: subject,
        body: emailBodyText,
        sender: sender,
      };
      // console.log("Sending reply suggestion request to Flask:", payload);

      const response = await fetch(SUGGEST_REPLY_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      // console.log("Flask reply suggestion response status:", response.status);
      if (response.ok) {
        const data = await response.json();
        // console.log("Flask reply suggestion data received:", data);
        if (data.suggested_replies && Array.isArray(data.suggested_replies)) {
          fetchedReplies = data.suggested_replies;
        } else {
          console.warn("Backend did not return an array of suggested replies in expected format.");
        }
      } else {
        const errorText = await response.text();
        console.error(
          `Error fetching suggested replies from backend (${response.status}): ${errorText}`
        );
      }
    } catch (error) {
      console.error("Network error fetching suggested replies:", error);
    }
    // console.log("Finished getSuggestedReplies. Found:", fetchedReplies.length, "replies.");
    return fetchedReplies;
  };

  const handleRefresh = () => {
    console.log("Refresh button clicked. Reloading email details...");
    loadEmailDetails(); // Re-run the main loading function
  };

  const handleReplyClick = (replyText) => {
    // console.log("Attempting to display reply form with text:", replyText.substring(0, 50) + "...");
    const replyHtml = replyText.replace(/\n/g, '<br />');
    const replyOptions = {
        htmlBody: replyHtml
    };
    Office.context.mailbox.item.displayReplyForm(replyOptions);
    // Office.context.mailbox.item.displayReplyForm(replyText);
  };


  // ----------------- NEW: Outgoing Email Validation Logic -----------------
  

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
        Refresh Content
      </button>

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
              checked={importanceEnabled}
              onToggle={(checked) => handleToggleChange("enable_importance", checked)}
            />
            <ToggleSwitch
              label="返信と要約の生成"
              checked={generationEnabled}
              onToggle={(checked) => handleToggleChange("enable_generation", checked)}
            />
          </Stack>

          {/* Analysis Result Section */}
          <Text variant="large" styles={{ root: { fontWeight: "bold" } }}>
            分析
          </Text>
          {importanceEnabled && (
            <>
              <Stack
                tokens={{ childrenGap: 10, padding: 15 }}
                styles={{ root: { border: "1px solid #ccccccff", borderRadius: 8 } }}
              >
                <FormattedText text={analysisResult} />
                {/* <Text variant="medium">{analysisResult}</Text> */}
              </Stack>
            </>
          )}
          {generationEnabled && (
            <>
              <SuggestedReplies replies={suggestedReplies} onReplyClick={handleReplyClick} />
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
