// Title: Change an Excel workbook's modify password using Aspose.Cells for .NET
// Description: Loads a password‑protected workbook, assigns a new write‑protection (modify) password via the Settings.WriteProtection property, and saves the file. The worksheet data, formulas, and formatting remain untouched.
// Keywords: Aspose.Cells change modify password | update Excel write protection .NET | replace workbook protection password C# | Aspose.Cells set new modify password | change Excel file password programmatically
// Common Searches: how to change modify password of an Excel file with Aspose.Cells | replace write‑protection password without altering workbook content | update Excel workbook protection password C# Aspose | change Excel file password programmatically .NET
// Developer Intent: Replace the existing modify (write‑protection) password of an Excel workbook while preserving all content and formatting.
// Use Cases: Rotate workbook modify passwords after a security policy change without re‑creating files. | Automate password updates for shared Excel reports before distribution. | Migrate legacy workbooks to a new corporate password standard while keeping data intact.
// AI Prompts: Generate C# code with Aspose.Cells that changes the modify password of an existing workbook without affecting its data. | Explain how to load an encrypted Excel file and update its write‑protection password using Aspose.Cells for .NET. | Show how to verify that only the modify password was changed while all worksheet content stays the same after saving.

using System;
using Aspose.Cells;

// Loads a password‑protected workbook, assigns a new write‑protection (modify) password via the Settings.WriteProtection property, and saves the file. The worksheet data, formulas, and formatting remain untouched.
class ChangeWriteProtectionPassword
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "ProtectedWorkbook.xlsx";

        // Existing write‑protection password (if the workbook is already protected)
        string currentPassword = "oldPassword";

        // New password that will replace the old one
        string newPassword = "newPassword";

        // Load the workbook with the current password (if any)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = currentPassword;          // for file encryption password, if set
        Workbook wb = new Workbook(inputPath, loadOptions);

        // Change the write‑protection (modify) password
        wb.Settings.WriteProtection.Password = newPassword;

        // Save the workbook – content remains unchanged, only the password is updated
        string outputPath = "Workbook_With_New_ModifyPassword.xlsx";
        wb.Save(outputPath);
    }
}
