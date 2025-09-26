# GEMINI.md

## Project Overview

This is a Google Apps Script project named "Spartan_Read_Aloud". It provides a "read aloud" functionality for assessments. The system is designed to take PDFs from a Google Sheet, process them to extract text, generate audio for the text chunks using the Gemini API, and then present the PDF and audio to students through a web interface.

The main components are:

*   **Google Sheet ("Assessment Database")**: This sheet acts as the database, holding links to PDF assessments, credentials, and the generated audio data.
*   **Google Apps Script Backend (`Code.js`, `Gemini.js`)**: This server-side logic, running on Google's servers, automates the processing of PDFs and generation of audio.
*   **Web App Frontend (`index.html`)**: A student-facing web application that allows students to log in, view the assessment PDF, and click on text to have it read aloud.

## Building and Running

This is a Google Apps Script project, so there are no traditional build or run commands. The project is deployed and run on Google's servers.

To work with the project locally, you would use `clasp`, the command-line interface for Google Apps Script.

*   **Pushing changes:** `clasp push`
*   **Pulling changes:** `clasp pull`
*   **Opening the project in the Apps Script editor:** `clasp open`

The web app is accessed via a URL provided by the Google Apps Script deployment.

## Development Conventions

*   The backend is written in JavaScript (Google Apps Script is based on JavaScript).
*   The frontend is a single HTML file with embedded CSS and JavaScript.
*   The project uses the Gemini API for text-to-speech. The API key is stored as a script property.
*   The project uses several Google Workspace APIs, including Drive, Sheets, and Documents. The required OAuth scopes are defined in `appsscript.json`.
*   The backend code is organized into two main files: `Code.js` for the main application logic and `Gemini.js` for the Gemini API interaction.
*   The code includes error handling and logging using `Logger.log()`.
