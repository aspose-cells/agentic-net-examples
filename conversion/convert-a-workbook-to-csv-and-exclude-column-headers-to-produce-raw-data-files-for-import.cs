// Title: Convert Excel to CSV without Headers using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook, deletes the first row of the first worksheet (assumed header), sets TxtSaveOptions.ExportAllSheets = true, and saves all sheets to a single CSV file that contains only raw data.
// Keywords: Aspose.Cells CSV export C# | remove header row Aspose.Cells | ExportAllSheets CSV Aspose | Excel to raw CSV .NET | Aspose.Cells TxtSaveOptions | C# convert Excel to CSV without headers | save multiple worksheets to one CSV
// Common Searches: Aspose.Cells export Excel to CSV without header row | C# save all worksheets as one CSV using Aspose.Cells | how to delete first row before CSV export Aspose.Cells | convert workbook to raw CSV Aspose.Cells .NET | export Excel data without column titles C#
// Developer Intent: Create a CSV file that contains only the data rows from an Excel workbook by removing column headings before export.
// Use Cases: Import raw data into a database where column titles are not required. | Combine data from several worksheets into a single CSV for batch processing. | Automate nightly ETL jobs that need header‑free CSV files.
// AI Prompts: Show C# code with Aspose.Cells that deletes the first row of each worksheet and exports the workbook to a single CSV file. | Explain how to use TxtSaveOptions.ExportAllSheets to write all Excel sheets to one CSV without headers. | Provide a modification to keep headers in selected sheets while removing them from others during CSV export.

using System;
using Aspose.Cells;

// Loads an .xlsx workbook, deletes the first row of the first worksheet (assumed header), sets TxtSaveOptions.ExportAllSheets = true, and saves all sheets to a single CSV file that contains only raw data.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file (raw data without headers)
        string csvPath = "output.csv";

        // Load the workbook from the source file
        Workbook workbook = new Workbook(sourcePath);

        // Remove the first row (assumed to contain column headers) from the first worksheet.
        // If you need to remove headers from all worksheets, iterate over workbook.Worksheets.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteRow(0);

        // Prepare CSV save options.
        // ExportAllSheets = true ensures that all worksheets are written to the CSV file.
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            ExportAllSheets = true
        };

        // Save the workbook as CSV using the specified options.
        workbook.Save(csvPath, saveOptions);

        Console.WriteLine($"Workbook converted to CSV without headers: {csvPath}");
    }
}
