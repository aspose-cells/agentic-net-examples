// Title: Change Worksheet Password While Keeping Protection Settings – Aspose.Cells for .NET (C#)
// Description: Load a workbook, retrieve a protected worksheet, and replace its password using the Protect method with the old password. All existing protection options (AllowEditing flags, protection type, etc.) remain unchanged. Save the workbook with the new password.
// Keywords: Aspose.Cells change worksheet password | C# protect worksheet preserve settings | update Excel sheet password programmatically | worksheet Protect method old new password | retain AllowEditing options Aspose
// Common Searches: how to change password of a protected worksheet in Aspose.Cells | replace worksheet password without losing protection options .NET | Aspose.Cells keep AllowEditing flags when updating password | C# change Excel sheet password programmatically preserving protection
// Developer Intent: Replace the password of a protected worksheet without altering any of its current protection configurations.
// Use Cases: Rotate worksheet passwords across multiple workbooks to comply with new security policies while preserving editing permissions. | Automate password renewal for protected sheets in a document management system without resetting protection flags. | Integrate password updates into a batch export process, ensuring the original protection type and allowances stay intact.
// AI Prompts: Show C# code that changes a worksheet password in Aspose.Cells while keeping all protection settings. | Explain how to use the Protect method to swap an existing worksheet password without resetting AllowEditing options. | Provide an example of updating the password of a protected Excel sheet with Aspose.Cells, preserving the current protection type.

using System;
using Aspose.Cells;

// Load a workbook, retrieve a protected worksheet, and replace its password using the Protect method with the old password. All existing protection options (AllowEditing flags, protection type, etc.) remain unchanged. Save the workbook with the new password.
class ChangeWorksheetPassword
{
    static void Main()
    {
        // Load the workbook that contains the protected worksheet
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the worksheet whose password you want to change
        Worksheet worksheet = workbook.Worksheets[0];

        // Existing password (must be known)
        string oldPassword = "oldPassword";

        // New password to set
        string newPassword = "newPassword";

        // Change the password while preserving all existing protection options.
        // The Protect method with the oldPassword parameter retains the current
        // protection settings (AllowEditing..., etc.).
        worksheet.Protect(ProtectionType.All, newPassword, oldPassword);

        // Save the workbook with the updated password
        workbook.Save("output.xlsx");
    }
}
