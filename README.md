# ESL Tutor Dashboard PWA

This GitHub Pages project is an installable PWA shell for the existing Google Apps Script ESL Tutor Dashboard.

## How it works

GitHub Pages hosts the PWA shell.
The shell opens the existing Apps Script web app inside it.
Google Sheets and Google Drive remain the existing backend/storage.

This avoids rewriting the current dashboard and avoids changing existing `google.script.run` functionality.

## First setup

Open the GitHub Pages site. It will ask for the deployed Apps Script Web App URL ending in `/exec`.

The URL is stored only in the browser's local storage.

## Important

The Apps Script deployment must allow the web app to be opened/embedded. The current project already uses `ALLOWALL` in `doGet()`.
