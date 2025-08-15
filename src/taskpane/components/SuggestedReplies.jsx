import React from 'react';
import { Stack, Text, PrimaryButton } from '@fluentui/react';

const SuggestedReplies = ({ replies, onReplyClick }) => {
  if (!replies || replies.length === 0) {
    return (
      <Stack tokens={{ childrenGap: 10, padding: 15 }} styles={{ root: { border: '1px solid #b2c0fcff', borderRadius: 8 } }}>
        <Text variant="large" styles={{ root: { fontWeight: 'bold' } }}>返信提案</Text>
        <Text variant="medium">No reply suggestions available.</Text>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 10, padding: 15 }} styles={{ root: { border: '1px solid #b2c0fcff', borderRadius: 8 } }}>
      <Text variant="large" styles={{ root: { fontWeight: 'bold' } }}>返信提案</Text>
      {replies.map((reply, index) => (
        <PrimaryButton
          key={index} // Unique key for list rendering
          // Truncate long reply texts for button labels for better UI
          text={reply.length > 70 ? `${reply.substring(0, 67)}...` : reply}
          onClick={() => onReplyClick(reply)} // Trigger the reply action
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
              // Added for rounded corners
              borderRadius: 8 ,
               backgroundColor: '#e8f8fdff',
               color: '#000000',
            } 
          }} // Ensure button content wraps and has rounded corners
        />
      ))}
    </Stack>
  );
};

export default SuggestedReplies;
