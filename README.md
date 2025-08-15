# Office Add-in Taskpane React JS

This project is an Outlook Add-in built with React. It provides a taskpane interface within Outlook.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/parvezamm3/mailai_addin.git
    ```
2.  **Navigate to the project directory:**
    ```bash
    cd outlook-addin-MailAIAddin
    ```
3.  **Install dependencies:**
    ```bash
    npm install
    ```

## Getting Started

1.  **Start the development server and sideload the add-in:**
    ```bash
    npm start
    ```
    This command will start the development server and automatically open Outlook with the add-in sideloaded.

2.  **Open the add-in in Outlook:**
    *   In Outlook, open an email.
    *   Click on the "MailAI" tab in the ribbon.
    *   Click on "Show Taskpane" to open the add-in.

## Available Scripts

*   `npm run lint`: Lints the code using ESLint.
*   `npm run lint:fix`: Fixes linting errors automatically.
*   `npm run validate`: Validates the `manifest.xml` file.
*   `npm run build`: Builds the add-in for production.

