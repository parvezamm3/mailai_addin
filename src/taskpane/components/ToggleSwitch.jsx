import React from 'react';
import { Toggle, Stack } from '@fluentui/react';

/**
 * Reusable ToggleSwitch component for boolean settings.
 * It uses Fluent UI's Toggle component.
 * @param {object} props - Component props.
 * @param {string} props.label - The text label for the toggle switch.
 * @param {boolean} props.checked - The current checked state of the toggle.
 * @param {(checked: boolean) => void} props.onToggle - Callback function when the toggle state changes.
 */
const ToggleSwitch = ({ label, checked, onToggle }) => {
    // console.log("Toggling Switch");
    // console.log(label, checked, onToggle);
  return (
    <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
      <Toggle
        label={label} // Displays the label text next to the toggle
        inlineLabel // Aligns the label inline with the toggle
        checked={checked} // Controls the checked state
        onChange={(event, checked) => onToggle(checked)} // Handles state changes
        // Fluent UI toggles have built-in styling for visual feedback
      />
    </Stack>
  );
};

export default ToggleSwitch;
