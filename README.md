# OutlookToExcelAttendance

## Overview
OutlookToExcelAttendance is a comprehensive tool designed to export attendance data from Microsoft Outlook to an Excel spreadsheet. This utility is ideal for professionals and organizations looking to efficiently manage and track attendance records.

## Key Features
- **Automated Data Export**: Extracts attendance information from Outlook and imports it directly into Excel.
- **Customizable Excel Sheet**: The Excel sheet is structured to display attendance data clearly, including attendee names, attendance status, response, and email addresses.
- **User-Friendly Interface**: Easy to use, catering to users with different levels of technical expertise.
- **VBA Scripting**: Utilizes Visual Basic for Applications (VBA) scripting to automate tasks within Excel.
- **Recurring Meeting Detection**: Includes a menu to select and import specific recurring meetings into the spreadsheet.

## Requirements
- **Microsoft Outlook**: Access to Outlook is necessary to retrieve attendance data.
- **Microsoft Excel**: Excel is used for data formatting and presentation.
- **VBA Enabled**: Ensure that VBA is enabled in Excel to run the scripts properly.
- **Local Execution**: The Excel file must be run locally on your computer and cannot be executed in cloud environments like Excel Online.

## Known Issues
- Special characters such as +, -, *, /, \ in the meeting title may result in unexpected outcomes.

## Excel Sheet Description
The Excel sheet is designed to display various details such as attendee names, their status (e.g., Meeting Organizer, Required Attendee, Optional Attendee), response status (Accepted, Tentative, Declined, None), and email addresses. It also includes summary columns with formulas to count the number of responses in each category.

## Included Files
- **VBA_MeetingStatus_Outlook_to_Excel_v1.2.1_beta.xlsm**: The Excel workbook containing the VBA scripts and structure for managing the exported data.

## Usage
1. **Setup**: Ensure that Microsoft Outlook and Excel are installed and that VBA is enabled in Excel.
2. **Open Excel File**: Open the `VBA_MeetingStatus_Outlook_to_Excel_v1.2.1_beta.xlsm` file.
3. **Enter Meeting Title**: In cell `K2` under the "Input parameters" section, enter the meeting title you want to retrieve data for.
4. **Retrieve Data**: Click the button labeled "Retrieve Data from Outlook" to import the attendance data into the Excel sheet.
5. **Send Email Reminders**: If needed, edit the email text in cell `K3` and click the button labeled "Send Reminder Email for None/Tentative Responses" to send out email reminders.

## Version History
- **v1.2.1 Beta**
  - Added detection for recurring meetings with a selection menu.
  - Fixed issue in email reminders to exclude attendees who declined the meeting.
  - Known bugs related to special characters in meeting titles.

## Authorization
Users must authorize the tool to access their Outlook data for security and data integrity.

## License
Refer to the license file for information on usage and redistribution.

## Contributing
Contributions are welcome. Please follow the contributing guidelines for more information on how to contribute.
