// Title: Remove Excel workbook password using Aspose.Cells for .NET
// Description: Load a password‑protected .xlsx file with LoadOptions.Password, call Workbook.Unprotect to clear the protection, and save the workbook as an unencrypted copy.
// Keywords: Aspose.Cells remove password | unprotect Excel workbook C# | load protected workbook Aspose | Workbook.Unprotect method | save unencrypted Excel file | C# Excel password removal
// Common Searches: how to remove password from Excel file using Aspose.Cells | Aspose.Cells C# unprotect workbook example | load protected .xlsx and save without password | programmatically delete Excel workbook password | Aspose.Cells remove workbook encryption
// Developer Intent: Open a password‑protected Excel workbook, strip its protection, and write an unprotected version to a new location.
// Use Cases: Batch conversion of secured reports to plain Excel files for downstream data pipelines. | Automating password removal before uploading workbooks to services that do not support Excel encryption. | Creating editable copies of protected workbooks for collaborators without sharing the original password.
// AI Prompts: Write C# code with Aspose.Cells that opens a protected .xlsx, removes the password, and saves the file to a given path. | Show how to catch and handle an invalid password exception when calling Workbook.Unprotect in Aspose.Cells. | Demonstrate unprotecting a workbook while preserving all formulas, formatting, and charts using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load a password‑protected .xlsx file with LoadOptions.Password, call Workbook.Unprotect to clear the protection, and save the workbook as an unencrypted copy.
class RemoveWorkbookPassword
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";

        // Password used to protect the workbook
        string password = "test";

        // Load the workbook with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = password;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Remove the workbook protection
        workbook.Unprotect(password);

        // Save the unprotected workbook to a new file
        string outputPath = "unprotected.xlsx";
        workbook.Save(outputPath);
    }
}
