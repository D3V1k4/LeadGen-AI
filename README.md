<div align="center">

<img src="assets/leadgen-ai-banner.png" alt="LeadGen AI Banner" width="100%">

# LeadGen AI

### AI-Powered Lead Management & Customer Engagement System

An intelligent automation solution that captures customer enquiries, analyses potential leads using Gemini AI and generates personalised email responses.

</div>

---

## About the Project

LeadGen AI is an AI-powered lead management system designed to simplify the process of handling customer enquiries.

Traditionally, businesses manually collect enquiries, evaluate potential customers and prepare responses. This project automates these repetitive tasks by connecting Google Forms, Google Sheets, Google Apps Script, Gemini AI and Gmail.

## Key Features

- **Automated Lead Collection:** Captures customer enquiries through Google Forms.
- **AI-Powered Analysis:** Uses Gemini AI to understand customer requirements.
- **Lead Scoring:** Assigns scores between 0 and 100.
- **Lead Prioritisation:** Categorises leads as High, Medium or Low priority.
- **AI-Generated Summaries:** Creates concise summaries of customer enquiries.
- **Personalised Email Generation:** Generates professional email subjects and responses.
- **Duplicate Lead Detection:** Identifies repeated submissions using email addresses.
- **Automated Data Management:** Stores lead information and AI-generated results in Google Sheets.
- **Error Handling:** Includes retry handling for temporary Gemini API server errors.

## Technology Stack

| Technology | Purpose |
|---|---|
| Google Forms | Lead collection |
| Google Sheets | Lead database |
| Google Apps Script | Workflow automation |
| Google Gemini API | AI-powered lead analysis |
| Gmail | Email delivery |
| JavaScript | Automation logic |

## System Workflow

1. Customer submits an enquiry through Google Forms.
2. The response is automatically stored in Google Sheets.
3. Google Apps Script detects the new submission.
4. The system checks whether the email address already exists.
5. Gemini AI analyses unique enquiries.
6. AI generates a summary, lead score, priority and personalised email.
7. Results are updated in Google Sheets.
8. Gmail sends the generated email to the configured test recipient.

## Lead Scoring

The system uses the following criteria to evaluate customer enquiries:

| Criteria | Maximum Points |
|---|---:|
| Budget suitability | 25 |
| Requirement clarity | 25 |
| Service relevance | 20 |
| Business potential | 20 |
| Readiness to proceed | 10 |
| **Total** | **100** |

Priority classification:

- High: 75–100
- Medium: 45–74
- Low: 0–44

## Project Structure

```text
LeadGen-AI/
│
├── assets/
│   └── leadgen-ai-banner.png
│
├── src/
│   └── Code.gs
│
└── README.md
```

## Setup Instructions

1. Create a Google Form containing customer enquiry fields.
2. Link the form to a Google Spreadsheet.
3. Open the spreadsheet and select Extensions → Apps Script.
4. Copy the source code from `src/Code.gs` into the Apps Script editor.
5. Open Project Settings → Script Properties.
6. Configure the following properties:

   - `GEMINI_API_KEY` – Your Gemini API key.
   - `TEST_EMAIL_ADDRESS` – Your email address for testing.

7. Ensure the response sheet is named `Form Responses 1`.
8. Run `setupAutomation()` and grant the required permissions.
9. Submit a sample enquiry through the Google Form.
10. Check the spreadsheet and test email inbox for the results.

**Security:** Never publish your API key or personal credentials in the repository. Configure them privately in Apps Script Script Properties.

## Current Project Status

The project has been implemented and tested using sample enquiries.

Email delivery currently operates in test mode, meaning generated emails are sent to a configured test recipient rather than directly to customers.

## Future Improvements

- Integration with CRM platforms.
- WhatsApp Business API integration.
- Lead analytics dashboard.
- Advanced lead performance tracking.
- Improved AI-based customer segmentation.


## Learning Outcomes

This project provided practical experience in API integration, AI-powered automation, JavaScript, data management, lead scoring and automated customer communication.

## Live Project Demo

You can explore the lead collection interface using the Google Form below:

[**Open LeadGen AI – Customer Enquiry Form**](https://forms.gle/TNbKcmfiAVE7JtRL7)

Note: This form is provided for demonstration purposes. Please use fictional sample information when submitting test enquiries.


---

<div align="center">

**Built as an AI Automation Internship Assessment Project**

</div>
