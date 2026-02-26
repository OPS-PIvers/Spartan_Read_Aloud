# GEMINI.md

## Project Overview

This is a Google Apps Script project named "Spartan Assessment Portal". It provides a "read aloud" functionality for assessments. The system is designed to take PDFs from a Google Sheet, process them to extract text, generate audio for the text chunks using the Gemini API, and then present the PDF and audio to students through a web interface.

The main components are:

*   **Google Sheet ("Assessment Database")**: This sheet acts as the database, holding links to PDF assessments, credentials, and the generated audio data.
*   **Google Apps Script Backend (`Code.js`, `Gemini.js`)**: This server-side logic, running on Google's servers, automates the processing of PDFs and generation of audio.
*   **Web App Frontend (`index.html`)**: A student-facing web application that allows students to log in, view the assessment PDF, and click on text to have it read aloud.

## Building and Running

This is a Google Apps Script project, so there are no traditional build or run commands. The project is deployed and run on Google's servers.

The web app is accessed via a URL provided by the Google Apps Script deployment.
