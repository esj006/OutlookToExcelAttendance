# OutlookToExcelAttendance

This GitHub repository contains VBA code for exporting participant lists and attendance status from Outlook meetings to an Excel workbook. This can be useful for keeping track of participants and their response statuses for meetings.

## Usage

The following steps provide an overview of how to use the VBA code:

1. Open the Excel workbook where you want to import data from the Outlook meeting.

2. Add the code to the VBA editor:
   - Copy and paste the VBA code from `EksporterDeltakereFraOutlookMote()` in this repository into the Excel VBA editor.

3. Modify the parameters:
   - Set the meeting title in Excel cell K2.
   - Optionally, customize the email text in Excel cell K3.

4. Run the code:
   - Execute the VBA code by clicking "Run" or pressing "F5" in the VBA editor.

5. Results:
   - The code will retrieve the participant list and attendance status from the Outlook meeting and display it in the Excel workbook.

## Features

- Retrieves participants based on the meeting title.
- Categorizes participants into meeting organizer, required attendees, and optional attendees.
- Displays participants' response statuses (Accepted, Tentative, Declined, No Response).
- Provides an overview of participants' email addresses.

## Requirements

- Microsoft Excel.
- Microsoft Outlook.

## Authorization

Please note that this solution may require authorization to access Outlook data. Ensure that you have the necessary permissions before using this code.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

**Note:** This is a simple README file. You can customize this file with additional information, usage instructions, and screenshots as needed to help users understand and use your VBA code.
