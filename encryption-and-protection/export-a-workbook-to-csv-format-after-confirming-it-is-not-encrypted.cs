// Title: C# – Export a Non‑Encrypted Excel Workbook to CSV with Aspose.Cells
// Description: Loads an Excel file, verifies that Workbook.Settings.IsEncrypted is false, then uses TxtSaveOptions (SaveFormat.Csv, ExportAllSheets = true) to save the workbook as a CSV file and releases resources.
// Keywords: Aspose.Cells CSV export C# | check workbook encryption Aspose | TxtSaveOptions SaveFormat.Csv | ExportAllSheets CSV | Workbook.Settings.IsEncrypted
// Common Searches: Aspose.Cells export to CSV after encryption check | C# convert Excel to CSV only if not encrypted | Save all worksheets as one CSV using Aspose.Cells | How to detect encrypted workbook with Aspose.Cells
// Developer Intent: Convert an Excel workbook to CSV only when the file is not password‑protected.
// Use Cases: Batch‑process a folder of .xlsx files, skipping encrypted ones, and generate a single CSV per workbook. | Create CSV reports from template workbooks that must remain unprotected. | Automate data pipelines that require CSV output but need to avoid decryption errors.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, checks Workbook.Settings.IsEncrypted, and saves it as CSV using TxtSaveOptions. | Explain error handling strategies for encrypted workbooks when converting to CSV with Aspose.Cells. | Show how to modify the example to produce separate CSV files for each worksheet while still checking encryption status.

using System;
using Aspose.Cells;

// Loads an Excel file, verifies that Workbook.Settings.IsEncrypted is false, then uses TxtSaveOptions (SaveFormat.Csv, ExportAllSheets = true) to save the workbook as a CSV file and releases resources.
class ExportWorkbookToCsv
{
    static void Main()
    {
        // Paths for input workbook and output CSV file
        string inputPath = "input.xlsx";
        string outputPath = "output.csv";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Verify that the workbook is not encrypted before exporting
        if (workbook.Settings.IsEncrypted)
        {
            Console.WriteLine("The workbook is encrypted and cannot be exported to CSV.");
        }
        else
        {
            // Create CSV save options; ExportAllSheets = true exports every worksheet
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                ExportAllSheets = true
            };

            // Save the workbook as CSV using the Save(string, SaveOptions) rule
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully exported to CSV at: {outputPath}");
        }

        // Clean up resources
        workbook.Dispose();
    }
}
