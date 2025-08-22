import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { loadTheme, createTheme, ThemeProvider } from '@fluentui/react';
import * as React from 'react';
import * as ReactDOM from 'react-dom'; // Keep this for now for ReactDOM.render fallback if needed
// Import createRoot from 'react-dom/client' for React 18+
import { createRoot } from 'react-dom/client'; 
import App from './components/App';

// Initialize Fluent UI icons (important for all Fluent UI components)
initializeIcons();
// console.log("Initializing");
/* global document, Office */

// Define the root DOM element where the React app will be mounted
const container = document.getElementById('container');
// Create a root for React 18. This is the new way to manage React component trees.
const root = createRoot(container);

// Function to render the React application
const render = (Component) => {
  root.render( // Use root.render instead of ReactDOM.render
    // ThemeProvider is crucial for Fluent UI components to pick up their styles
    <ThemeProvider>
      <Component />
    </ThemeProvider>
  );
};

// Check if Office.js is ready before rendering the app
Office.onReady(() => {
  // We no longer need the isOfficeInitialized state, as render() will be called only once Office.js is ready.
  // console.log("Office is Ready");
  render(App);
});

// Removed the react-hot-loader (HMR) specific code block.
// Hot Module Replacement is typically for development convenience and
// can be removed if causing "Module not found" errors or for production builds.
