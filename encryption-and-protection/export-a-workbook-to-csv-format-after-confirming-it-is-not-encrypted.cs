// Title: Export an Unencrypted Excel Workbook to CSV with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file using Aspose.Cells, checks Workbook.Settings.IsEncrypted, and if the workbook is not password‑protected saves all worksheets to a single CSV via TxtSaveOptions (SaveFormat.Csv, ExportAllSheets = true).
// Keywords: Aspose.Cells | C# | .NET | export to CSV | unencrypted workbook | check encryption | TxtSaveOptions | SaveFormat.Csv | ExportAllSheets | Excel to CSV conversion | password protected Excel
// Common Searches: Aspose.Cells export workbook to CSV C# | How to check if Excel file is encrypted with Aspose.Cells | Save all sheets as CSV using Aspose.Cells .NET | Convert unencrypted .xlsx to CSV Aspose | C# code sample export Excel to CSV after encryption check
// Developer Intent: Convert a non‑encrypted Excel file to CSV using Aspose.Cells.
// Use Cases: Automated batch conversion of uploaded .xlsx files to CSV for data pipelines. | Server‑side generation of CSV reports from user‑provided Excel workbooks after confirming they are not password‑protected. | Integrating CSV export into a .NET application where encryption status must be validated first. | Creating a single CSV that consolidates all worksheets for downstream analytics.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, verifies Workbook.Settings.IsEncrypted is false, and saves the file as CSV with all sheets included. | Show how to handle a workbook that is encrypted by displaying an error message and skipping the export using Aspose.Cells. | Provide an example of setting a custom delimiter and encoding for CSV export with TxtSaveOptions in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an .xlsx file using Aspose.Cells, checks Workbook.Settings.IsEncrypted, and if the workbook is not password‑protected saves all worksheets to a single CSV via TxtSaveOptions (SaveFormat.Csv, ExportAllSheets = true).
class ExportToCsv
{
    static void Main()
    {
        // Paths for input workbook and output CSV file
        string inputPath = "input.xlsx";
        string outputPath = "output.csv";

        // Load the workbook from file
        Workbook workbook = new Workbook(inputPath);

        // Verify that the workbook is not encrypted
        if (workbook.Settings.IsEncrypted)
        {
            Console.WriteLine("The workbook is encrypted and cannot be exported.");
            return;
        }

        // Create CSV save options and export all worksheets
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.ExportAllSheets = true; // export every sheet, similar to Excel's behavior

        // Save the workbook as CSV using the provided Save method
        workbook.Save(outputPath, csvOptions);

        Console.WriteLine("Workbook successfully exported to CSV.");
    }
}
