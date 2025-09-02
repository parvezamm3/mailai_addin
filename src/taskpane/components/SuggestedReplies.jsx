import React from 'react';
import { Stack, Text, PrimaryButton } from '@fluentui/react';

const buttonStyle = { 
    width: '100%', 
    wordBreak: 'break-word', 
    whiteSpace: 'normal', 
    height: 'auto', 
    minHeight: '32px',
    padding: '5px',
    fontSize: 12,
    fontWeight: 100,
    borderRadius: 8,
    backgroundColor: '#e8f8fdff',
    color: '#000000',
};
const keyLabels = {
    Concise: '簡潔',
    Confirm: '確認',
    Polite: '丁寧',
};
const SuggestedReplies = ({ replies, onReplyClick }) => {
  // console.log(replies);
  return (
    <Stack tokens={{ childrenGap: 10, padding: 15 }} styles={{ root: { border: '1px solid #b2c0fcff', borderRadius: 8 } }}>
      <Text variant="large" styles={{ root: { fontWeight: 'bold' } }}>返信提案</Text>
      
      {replies && Object.keys(replies).length > 0 ? (
        Object.entries(replies).map(([key, value], index) => {
                // Get the Japanese label from the keyLabels map
                const label = keyLabels[key];
                
                // Truncate the message to the first 100 characters
                const truncatedMessage = value.length > 100 ? `${value.substring(0, 100)}...` : value;
                
                // Combine the label and the truncated message
                const buttonText = `${label} | ${truncatedMessage}`;

                return (
                    <PrimaryButton
                        key={index}
                        text={buttonText}
                        onClick={() => onReplyClick(value)}
                        styles={{ 
                            root: buttonStyle
                        }}
                    />
                );
            })
      ) : (
        <Text variant="medium">返信の提案はありません。</Text>
      )}
    </Stack>
  );
};

export default SuggestedReplies;