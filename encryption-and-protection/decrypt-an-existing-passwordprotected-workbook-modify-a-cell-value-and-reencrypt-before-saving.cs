// Title: C# – Open, Edit, and Re‑encrypt a Password‑Protected Excel Workbook with Aspose.Cells
// Description: Shows how to load a password‑protected .xlsx using Aspose.Cells LoadOptions, change a cell value, assign a new password via Workbook.Settings.Password, and save the workbook encrypted.
// Keywords: Aspose.Cells | C# password protected Excel | load encrypted workbook | modify cell Aspose.Cells | re‑encrypt workbook | Workbook.Settings.Password | LoadOptions.Password | Excel encryption .NET | code example | GitHub sample
// Common Searches: open password protected Excel file Aspose.Cells C# | change cell value in encrypted workbook Aspose | re‑encrypt Excel workbook after editing .NET | load and save protected XLSX with different passwords | Aspose.Cells example for workbook encryption
// Developer Intent: Load a protected Excel file, update a cell, and save it with a (new) password using Aspose.Cells for .NET.
// Use Cases: Automate updates to confidential reports stored in encrypted Excel files while preserving security. | Rotate workbook passwords after data corrections or periodic policy changes. | Batch‑process multiple protected workbooks to apply a standard password after modifications.
// AI Prompts: Write C# code with Aspose.Cells to open an encrypted .xlsx, change cell C5, and save it using a different password. | Explain the encryption workflow in Aspose.Cells and which properties control opening and saving passwords. | Provide a try‑catch example that handles an invalid opening password when loading a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load a password‑protected .xlsx using Aspose.Cells LoadOptions, change a cell value, assign a new password via Workbook.Settings.Password, and save the workbook encrypted.
class Program
{
    static void Main()
    {
        // Path to the existing password‑protected workbook
        string inputFile = "protected_workbook.xlsx";

        // Password required to open the workbook
        string openPassword = "oldPassword";

        // Password to apply after modification (can be same as openPassword)
        string newPassword = "newPassword";

        // Load the workbook with the opening password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = openPassword;
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Modify a cell value (example: cell B2 on the first worksheet)
        workbook.Worksheets[0].Cells["B2"].PutValue("Updated value");

        // Re‑encrypt the workbook with the new password
        workbook.Settings.Password = newPassword;

        // Save the modified and re‑encrypted workbook
        string outputFile = "modified_protected_workbook.xlsx";
        workbook.Save(outputFile);
    }
}
