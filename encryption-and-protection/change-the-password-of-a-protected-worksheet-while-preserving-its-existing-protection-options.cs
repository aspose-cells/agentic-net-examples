// Title: Change Worksheet Password While Preserving Protection Options – Aspose.Cells for .NET (C#)
// Description: Load a workbook, access a protected worksheet, and use the Protect method overload (ProtectionType.All, newPassword, oldPassword) to replace the worksheet password without altering any existing protection settings, then save the updated file.
// Keywords: Aspose.Cells change worksheet password | C# protect worksheet without losing settings | update worksheet password .NET | Preserve worksheet protection options | Protect method overload Aspose.Cells | worksheet password replacement | Aspose.Cells encryption and protection
// Common Searches: how to change password of a protected worksheet Aspose.Cells | replace worksheet password keep protection settings | Aspose.Cells Protect overload old password new password | C# change worksheet password without losing protection | update worksheet password programmatically Aspose
// Developer Intent: Replace the password of an already protected worksheet while leaving all current protection flags unchanged.
// Use Cases: Rotate worksheet passwords for compliance audits without unlocking cells, rows, or columns. | Share a workbook with a new team by assigning a new password but keeping the original protection layout. | Automate bulk password updates across many protected worksheets in a server‑side .NET process.
// AI Prompts: Provide C# code that changes a protected worksheet's password using Aspose.Cells while retaining all protection options. | Explain the Protect method overload that accepts old and new passwords for a worksheet in Aspose.Cells. | Generate a script to batch‑update worksheet passwords in multiple Excel files with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load a workbook, access a protected worksheet, and use the Protect method overload (ProtectionType.All, newPassword, oldPassword) to replace the worksheet password without altering any existing protection settings, then save the updated file.
class ChangeWorksheetPassword
{
    static void Main()
    {
        // Load the workbook that contains a protected worksheet
        Workbook workbook = new Workbook("ProtectedWorkbook.xlsx");

        // Access the worksheet whose password you want to change
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the current (old) password and the new password
        string oldPassword = "oldPass";
        string newPassword = "newPass";

        // Change the password while keeping all existing protection options intact
        // The third parameter supplies the old password required for the change
        worksheet.Protect(ProtectionType.All, newPassword, oldPassword);

        // Save the workbook with the updated password
        workbook.Save("ProtectedWorkbook_Updated.xlsx");
    }
}
