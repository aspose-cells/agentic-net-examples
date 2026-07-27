// Title: Set and Remove an Opening Password on an Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to assign an opening password to a workbook via Workbook.Settings.Password, save the protected file, reload it using LoadOptions.Password, clear the password, and save an unprotected copy while keeping all worksheet data intact.
// Keywords: Aspose.Cells C# | set workbook opening password | remove Excel password programmatically | load password‑protected workbook .NET | Workbook.Settings.Password | LoadOptions.Password | Excel encryption Aspose | unprotect Excel file C# | protect Excel workbook with Aspose.Cells | Aspose.Cells encryption and decryption
// Common Searches: How to add an opening password to an Excel file using Aspose.Cells | Remove opening password from a workbook with Aspose.Cells C# | Load a password‑protected Excel workbook in .NET | Save an unprotected copy of a protected Excel file using Aspose.Cells | Aspose.Cells example for workbook encryption and decryption
// Developer Intent: Apply an opening password to a workbook, then remove it and save the file without protection.
// Use Cases: Generate a report, protect it with a temporary password for secure transmission, and later store an unprotected version for archiving. | Open a password‑protected template, programmatically modify its content, and save a clean workbook for downstream processing. | Automate a batch workflow where files are initially encrypted for compliance and subsequently released as plain Excel files.
// AI Prompts: Show C# code using Aspose.Cells to set an opening password, load the workbook with LoadOptions, clear the password, and save an unprotected file. | Explain how Workbook.Settings.Password differs from WorkbookProtection and how to remove an opening password in Aspose.Cells. | Provide a step‑by‑step guide for encrypting an Excel workbook, opening it with a password, and then removing the password while preserving all data.

using System;
using Aspose.Cells;

// Demonstrates how to assign an opening password to a workbook via Workbook.Settings.Password, save the protected file, reload it using LoadOptions.Password, clear the password, and save an unprotected copy while keeping all worksheet data intact.
class Program
{
    static void Main()
    {
        // ---------- Create a new workbook and add sample data ----------
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // ---------- Apply an opening password ----------
        workbook.Settings.Password = "OpenPassword123";

        // Save the password‑protected workbook
        string protectedFile = "protected.xlsx";
        workbook.Save(protectedFile);

        // ---------- Load the protected workbook using the password ----------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "OpenPassword123";
        Workbook loadedWorkbook = new Workbook(protectedFile, loadOptions);

        // ---------- Remove the opening password ----------
        loadedWorkbook.Settings.Password = null; // clearing the password removes protection

        // Save the workbook without a password (unchanged content)
        string unprotectedFile = "unprotected.xlsx";
        loadedWorkbook.Save(unprotectedFile);
    }
}
