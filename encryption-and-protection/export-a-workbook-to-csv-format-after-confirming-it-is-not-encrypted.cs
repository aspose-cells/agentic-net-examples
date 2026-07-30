// Title: Export Unencrypted Excel Workbook to CSV with Aspose.Cells for .NET
// Description: Load an .xlsx file using Aspose.Cells, verify that the workbook is not password‑protected via the Settings.IsEncrypted property, and then save the active sheet as a CSV file. If the file is encrypted, the code reports the condition without attempting conversion.
// Keywords: Aspose.Cells CSV export | C# workbook encryption check | IsEncrypted property | SaveFormat.Csv .NET | convert Excel to CSV without password | Aspose.Cells .NET file conversion | Excel to CSV programmatically
// Common Searches: Aspose.Cells export to CSV only if not encrypted | Check workbook encryption before saving as CSV C# | How to skip password‑protected Excel files when converting to CSV with Aspose.Cells | C# code to detect encrypted workbook Aspose.Cells | Save active worksheet as CSV using Aspose.Cells .NET
// Developer Intent: Convert an Excel workbook to CSV only when it is not password‑protected.
// Use Cases: Process user‑uploaded spreadsheets in a web service, converting clear files to CSV while rejecting encrypted ones. | Create nightly CSV extracts from internal reports, automatically ignoring workbooks that require a password. | Build a batch script that scans a directory of .xlsx files and generates CSV versions for non‑encrypted items.
// AI Prompts: Generate C# code that scans a folder, opens each .xlsx with Aspose.Cells, checks Settings.IsEncrypted, and writes a .csv for files that are not encrypted. | Write a reusable method `bool TryExportToCsv(string xlsxPath, string csvPath)` that returns false when the workbook is encrypted and true after successful CSV creation, including exception handling. | Provide a PowerShell script that invokes a .NET assembly using Aspose.Cells to perform the same encryption check and CSV export.

using System;
using Aspose.Cells;

// Load an .xlsx file using Aspose.Cells, verify that the workbook is not password‑protected via the Settings.IsEncrypted property, and then save the active sheet as a CSV file. If the file is encrypted, the code reports the condition without attempting conversion.
class ExportWorkbookToCsv
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Check if the workbook is encrypted (requires a password to open)
        if (!workbook.Settings.IsEncrypted)
        {
            // Export to CSV format (active sheet)
            workbook.Save("output.csv", SaveFormat.Csv);
            Console.WriteLine("Workbook exported successfully to output.csv");
        }
        else
        {
            Console.WriteLine("The workbook is encrypted and cannot be exported without a password.");
        }

        // Clean up
        workbook.Dispose();
    }
}
