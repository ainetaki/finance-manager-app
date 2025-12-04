# finance-manager-app
A personal finance tracker built with Google Apps Script

Finance Manager (Google Apps Script Web App)
A modern, lightweight, and responsive personal finance tracker that runs entirely on Google Sheets.

This application is built as a Single Page Application (SPA) using Google Apps Script, HTML5, and Tailwind CSS. It provides a clean, app-like interface for managing your monthly bills, tracking income, and visualizing your annual financial health without needing a dedicated server or subscription service.



✨ Key Features
📊 Interactive Dashboard: Annual overview with charts (Chart.js) showing income, expenses, and surplus.
<img width="1900" height="1022" alt="image" src="https://github.com/user-attachments/assets/b7ecb86d-0195-44f6-adad-23bf196c52b0" />


📅 Monthly Management: Easily add, edit, and delete bills, income, and extra expenses.
<img width="480" height="631" alt="image" src="https://github.com/user-attachments/assets/edbec448-5569-45ed-a730-6f2754f53e60" />


🖱️ Drag & Drop Interface: Move bills between "Unpaid" and "Paid" columns simply by dragging them.
<img width="1904" height="1014" alt="image" src="https://github.com/user-attachments/assets/3cebc747-06f1-40e5-86f1-9f6560c79784" />


🔄 Recurring Transactions: Support for recurring bills (e.g., monthly, specific months, or the whole year).

🌍 Multi-language Support: Fully localized in English, Finnish, and Swedish.

<img width="429" height="709" alt="image" src="https://github.com/user-attachments/assets/45b1ee49-b7d0-4d69-a798-fc5064708941" />


📱 Responsive Design: Works perfectly on both desktop and mobile devices.

🔒 Secure: Simple password protection (stored safely in Script Properties) to prevent unauthorized access.

💾 Auto-Save: All data is saved directly to your Google Sheet in a structured format.


🛠️ Tech Stack:
- Backend: Google Apps Script (JavaScript)
- Frontend: HTML5, CSS3
- Styling: Tailwind CSS (via CDN)
- Charts: Chart.js (via CDN)
- Icons: FontAwesome (via CDN)
- Database: Google Sheets
  


🚀 Installation Guide
Since this app runs on Google's infrastructure, you don't need to install Node.js or set up a server. Just follow these steps:

1. Create the Google Sheet -> Go to Google Sheets and create a new, empty spreadsheet.

Name it whatever you like (e.g., "My Finance 2025").

2. Add the Code
In the spreadsheet, go to Extensions > Apps Script.

Delete any code in the default Code.gs file.

Copy the contents of Code.gs from this repository and paste it into the editor.

Click the + icon next to "Files" and select HTML. Name the file index.

Copy the contents of index.html from this repository and paste it into the new file.

Save the project (Ctrl+S or Cmd+S).

3. Deploy the Web App
In the Apps Script editor, click Deploy (blue button) > New deployment.

Click the gear icon (Select type) next to "Select type" and choose Web app.

Fill in the details:

Description: Finance Manager v1

Execute as: Me (This is important!)

Who has access: Only myself (Recommended for personal use) or Anyone with Google account (If you want to share it).

Click Deploy.

Authorize Access: Google will ask for permission to access your spreadsheets. Click "Review Permissions", select your account, and if you see a warning screen ("Google hasn't verified this app"), click Advanced > Go to (Untitled Project) (unsafe). This is safe because you are the author of the code.

4. Setup Wizard
Copy the Web App URL provided after deployment.

Open the URL in your browser.

You will see a Setup Wizard. Enter:

The starting year (e.g., 2025).

Your preferred language and currency.

Create a login password.

Click Start Using App.

<img width="452" height="657" alt="image" src="https://github.com/user-attachments/assets/1037b414-d2f6-4e7f-bd50-ffaf8104aecf" />


📖 How to Use
Dashboard: Use the top navigation arrows to switch between the Annual View (Year) and Monthly View.

Adding Bills: Click the floating + button to add a new bill, income, or other expense.

Tip: Select "12 Months (Year)" recurrence to add a bill to every month of the remaining year.

Paying Bills: In the monthly view, drag a bill card from the Unpaid (Red) column to the Paid (Green) column. The totals update automatically.

Settings: Click the gear icon to change language, currency, update categories, or change your password.

📄 License
This project is open-source and available under the MIT License. You are free to use, modify, and distribute it.

🔐 Security & Architecture

This application was built with a "Privacy First" architecture. Unlike many finance apps that store your data on external third-party servers, Finance Manager runs entirely within your personal Google ecosystem.

Here is how the security is implemented technically:

1. Data Sovereignty (Google Sheets)
No External Database: All financial data is stored in your personal Google Sheet. No data is ever sent to an external server, analytics platform, or third-party database.
Ownership: You are the sole owner of the data. If you delete the spreadsheet, the data is gone forever. No copies exist elsewhere.

2. Server-Side Authentication
Secure Password Storage: The application uses a custom password protection system. The password you create during setup is not stored in the HTML code or JavaScript files where it could be seen by "Viewing Source".
- Implementation: The password is saved using Google Apps Script's PropertiesService.getScriptProperties(). This is a secure, server-side key-value store that is inaccessible to the client-side browser.
- Verification: When you log in, your input is sent to the backend (Google Server). The comparison happens on the server, and only a boolean true/false is returned to the browser.

3. Execution Context
"Execute as Me": The Web App is deployed to run under your authority (Execute as: Me). This means the script relies on your Google Account's existing security (including your 2-Factor Authentication if enabled on your Google Account).
Access Control: By setting the deployment access to Only myself, you ensure that even if someone guesses your Web App URL, they cannot access the application without being logged into your Google Account and knowing the app-specific password.

4. Code Transparency
Open Source: Since you host the code yourself within the Google Sheet's script editor, you have full visibility. There is no compiled or obfuscated code—what you see is exactly what runs.
