// Title: Remove Excel Modify Password but Keep Opening Password with Aspose.Cells for .NET (C#)
// Description: Loads a workbook using its opening (encryption) password, clears the write‑protection (modify) password via Workbook.Settings.WriteProtection.Password, and saves the file so only the opening password remains.
// Keywords: Aspose.Cells | C# | remove modify password | clear write protection | keep opening password | Excel encryption | Workbook.Settings.WriteProtection | Excel file protection | Aspose.Cells .NET | delete write‑protection password
// Common Searches: Aspose.Cells remove modify password C# | How to keep opening password after deleting write protection in Excel | Clear workbook modify password using Aspose.Cells .NET | Remove write‑protection password while preserving encryption password | C# code to delete Excel modify password with Aspose
// Developer Intent: Delete the workbook’s modify/write‑protection password while preserving its opening (encryption) password using Aspose.Cells for .NET.
// Use Cases: Load a password‑protected Excel file, strip the modify password, and save it so users only need the opening password to view the workbook. | Prepare distribution‑ready workbooks by removing edit restrictions while maintaining encryption against unauthorized access. | Automate batch processing that clears modify passwords from multiple files without altering their original opening passwords.
// AI Prompts: Provide C# code with Aspose.Cells that removes the write‑protection password from an Excel workbook but retains the opening password. | Generate a reusable method that accepts a file path and opening password, loads the workbook, clears the modify password, and saves the result. | Explain how Workbook.Settings.WriteProtection.Password can be set to null or empty to delete a modify password in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads a workbook using its opening (encryption) password, clears the write‑protection (modify) password via Workbook.Settings.WriteProtection.Password, and saves the file so only the opening password remains.
class RemoveModifyPassword
{
    static void Main()
    {
        // Opening (encryption) password of the workbook
        string openingPassword = "openPwd";

        // Load the workbook using the opening password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = openingPassword;
        Workbook workbook = new Workbook("protected.xlsx", loadOptions);

        // Preserve the opening password for saving
        workbook.Settings.Password = openingPassword;

        // Remove the "password to modify" (write‑protection password)
        workbook.Settings.WriteProtection.Password = null; // or string.Empty

        // Save the workbook; it will remain protected with the opening password only
        workbook.Save("unprotected_modify.xlsx");
    }
}
