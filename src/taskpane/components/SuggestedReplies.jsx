import React from 'react';
import { Stack, Text, PrimaryButton } from '@fluentui/react';

const SuggestedReplies = ({ replies, onReplyClick }) => {
  return (
    <Stack tokens={{ childrenGap: 10, padding: 15 }} styles={{ root: { border: '1px solid #b2c0fcff', borderRadius: 8 } }}>
      <Text variant="large" styles={{ root: { fontWeight: 'bold' } }}>返信提案</Text>
      {/* Use a ternary operator for conditional rendering */}
      {replies && replies.length > 0 ? (
        // Map over the replies array if it exists and is not empty
        replies.map((reply, index) => (
          <PrimaryButton
            key={index}
            text={reply.length > 70 ? `${reply.substring(0, 67)}...` : reply}
            onClick={() => onReplyClick(reply)}
            styles={{ 
              root: { 
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
              } 
            }}
          />
        ))
      ) : (
        // Display a message if no replies are available
        <Text variant="medium">返信の提案はありません。</Text>
      )}
    </Stack>
  );
};

export default SuggestedReplies;