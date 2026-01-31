# AI Exam Grader (Apps Script + Groq)

A powerful, automated exam grading system that uses Google Apps Script and the Groq API (Llama 3.3 70B) to grade student responses based on specific criteria.

## Features

- **Automated Grading**: Automatically triggers on Google Form submissions.
- **Criteria-Based Scoring**: Grades each question according to custom JSON criteria.
- **Batch Processing**: Groups questions into chunks to optimize API usage and speed.
- **Robust Parsing**: Advanced JSON extraction logic handles variations in AI responses.
- **Email Feedback**: Sends a professional HTML email to students with their scores, feedback, and improvement tips.
- **Sheet Integration**: Logs everything to an `AI_GRADES` sheet for easy tracking.

## Setup Instructions

### 1. Spreadsheet Preparation
Your Master Spreadsheet must have the following sheets:

- **CONFIG**: 
  - Columns: `response_spreadsheet_id`, `response_sheet_name`, `email_column_header`, `exam_name`, `hf_model`, `name_column_header`.
- **ANSWER_KEY**:
  - Columns: `question_header` (exact question text from form), `question_id` (e.g., Q1), `criteria_json`.
- **AI_GRADES**: (Automatically created if it doesn't exist)
  - Logs timestamp, email, exam name, scores, and full AI output.

### 2. Apps Script Setup
1. Open your Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Paste the contents of `main.js` into the editor.
4. Save the project.

### 3. Configuration
Set the following in **Project Settings > Script Properties**:

- `GROQ_API_KEY`: Your API key from [Groq Console](https://console.groq.com/).
- `MASTER_SPREADSHEET_ID`: The ID of your master spreadsheet (found in the URL).

### 4. Install Triggers
In the Apps Script editor, run the `installTriggers()` function once. This will set up the `onFormSubmit` trigger for all configured response spreadsheets.

## File Overview

- `main.js`: The core logic for grading, parsing, and emailing.
- `questions.txt`: Sample questions and criteria formats.
- `.env`: Template for local environment variables.
- `.gitignore`: Ensures secrets like your API key stay local.

## Troubleshooting

- **Missing Questions in Email**: Ensure `question_id` in `ANSWER_KEY` matches the expected format.
- **API Errors**: Check the Apps Script execution logs to see the `Groq API response code`. 
- **Parsing Failures**: The script includes a robust fallback parser, but ensure your `criteria_json` is valid JSON.

## License
Created by Yared Technology School.
