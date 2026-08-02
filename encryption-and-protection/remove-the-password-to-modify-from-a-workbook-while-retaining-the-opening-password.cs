// Title: C# – Remove Excel ‘modify’ password while preserving the open password using Aspose.Cells
// Description: Load a password‑protected workbook with Aspose.Cells, clear the write‑protection password by setting Workbook.Settings.WriteProtection.Password to null, and save the file so the opening password remains unchanged.
// Keywords: Aspose.Cells remove modify password | C# clear write protection Excel | keep opening password Aspose | Workbook.Settings.WriteProtection.Password | .NET Excel encryption | remove write‑protect password programmatically
// Common Searches: how to delete modify password from Excel with Aspose.Cells | retain open password after removing write protection .NET | Aspose.Cells clear workbook modify password | remove Excel write‑protect password without affecting open password
// Developer Intent: Strip the workbook’s modify password while leaving the opening password intact.
// Use Cases: Automate batch cleaning of Excel templates that should open with a known password but be freely editable. | Prepare a shared workbook where users only need the open password, not a separate edit password. | Integrate into a CI pipeline to ensure Excel files are not blocked by write‑protect passwords.
// AI Prompts: Provide C# code with Aspose.Cells that removes the modify password from an Excel file but keeps the open password. | Write a method that accepts a file path and opening password, clears the write‑protection password, and saves a new copy. | Explain why setting Workbook.Settings.WriteProtection.Password to null removes the modify password in Aspose.Cells.

using System;
using Aspose.Cells;

// Load a password‑protected workbook with Aspose.Cells, clear the write‑protection password by setting Workbook.Settings.WriteProtection.Password to null, and save the file so the opening password remains unchanged.
class RemoveWriteProtection
{
    static void Main()
    {
        // Specify the password required to open the workbook
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "openPwd";

        // Load the workbook using the opening password
        Workbook workbook = new Workbook("protected.xlsx", loadOptions);

        // Clear the password that protects the file from being modified (write protection)
        // Setting it to null (or string.Empty) removes the "password to modify" while keeping the opening password intact
        workbook.Settings.WriteProtection.Password = null;

        // Save the workbook; the opening password remains unchanged
        workbook.Save("unprotected_modify.xlsx");
    }
}
