// Title: How to Remove Password Protection from an Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Load a password‑protected workbook with LoadOptions, call Unprotect on the workbook and each worksheet, clear the encryption password in Workbook.Settings, and save the file as an unprotected Excel document.
// Keywords: Aspose.Cells unprotect workbook | remove Excel password C# | clear worksheet protection Aspose | load encrypted workbook .NET | save unprotected Excel file
// Common Searches: Aspose.Cells remove workbook password | C# unprotect Excel file with Aspose | how to delete worksheet protection using Aspose.Cells | strip encryption from Excel workbook .NET | unprotect protected.xlsx Aspose.Cells
// Developer Intent: Load a protected Excel file, strip all workbook and worksheet passwords, and write a new file that can be opened and edited without any authentication.
// Use Cases: Automated preprocessing of client‑submitted spreadsheets before data‑analysis pipelines. | Preparing reports for distribution where password barriers are not allowed. | Compliance‑driven sanitization of uploaded Excel files by removing all protection layers.
// AI Prompts: Write C# code with Aspose.Cells that opens a password‑protected workbook, removes workbook and sheet passwords, and saves an unprotected copy. | Explain the purpose of Workbook.Settings.Password and how to clear it after unprotecting a file. | Show error‑handling patterns for incorrect passwords when loading a protected workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Load a password‑protected workbook with LoadOptions, call Unprotect on the workbook and each worksheet, clear the encryption password in Workbook.Settings, and save the file as an unprotected Excel document.
class RemovePasswordProtection
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";

        // Password that protects the workbook and its worksheets
        string password = "myPassword";

        // Load the workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Remove workbook structure/window protection
        workbook.Unprotect(password);

        // Remove protection from every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Unprotect(password);
        }

        // Clear any encryption password stored in the workbook settings
        workbook.Settings.Password = null;

        // Save the workbook without any password protection
        string outputPath = "unprotected.xlsx";
        workbook.Save(outputPath);
    }
}
