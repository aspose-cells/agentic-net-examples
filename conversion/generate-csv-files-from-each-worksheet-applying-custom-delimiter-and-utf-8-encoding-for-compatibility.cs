// Title: Export Excel worksheets to separate CSV files with a custom delimiter and UTF‑8 encoding using Aspose.Cells for .NET
// Description: Load a workbook, set TxtSaveOptions to CSV, define a custom separator (e.g., ';'), enable UTF‑8 encoding, and turn on ExportAllSheets so Aspose.Cells creates one CSV file per worksheet. Ideal for batch conversions and locale‑specific CSV output.
// Keywords: Aspose.Cells CSV export | custom delimiter CSV C# | UTF-8 CSV Aspose.Cells | ExportAllSheets | save each sheet as CSV | .NET Excel to CSV conversion | TxtSaveOptions separator
// Common Searches: Aspose.Cells export each worksheet to CSV | C# set CSV delimiter semicolon Aspose.Cells | UTF-8 CSV output from Excel using Aspose | how to generate multiple CSV files from one workbook | ExportAllSheets CSV Aspose.Cells example
// Developer Intent: Create individual CSV files for every worksheet in an Excel workbook, using a specified field separator and UTF‑8 encoding.
// Use Cases: Produce locale‑specific CSV reports where a semicolon or other delimiter is required. | Generate UTF‑8 encoded CSV files for downstream systems that only accept UTF‑8 data. | Automate batch conversion of multi‑sheet Excel workbooks into separate CSV files for data pipelines.
// AI Prompts: Write C# code with Aspose.Cells to export each worksheet to CSV using a pipe (|) delimiter and UTF‑8 encoding. | Explain the effect of the ExportAllSheets property in TxtSaveOptions when saving to CSV. | Show how to specify an output folder and custom file‑naming pattern for the CSV files created per worksheet.

using System;
using System.Text;
using Aspose.Cells;

// Load a workbook, set TxtSaveOptions to CSV, define a custom separator (e.g., ';'), enable UTF‑8 encoding, and turn on ExportAllSheets so Aspose.Cells creates one CSV file per worksheet. Ideal for batch conversions and locale‑specific CSV output.
class ExportWorksheetsToCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ';' ;            // custom delimiter
        csvOptions.Encoding = Encoding.UTF8;    // UTF‑8 encoding
        csvOptions.ExportAllSheets = true;      // export each worksheet

        // Save the workbook; Aspose.Cells will create a separate CSV file for each sheet
        workbook.Save("output.csv", csvOptions);
    }
}
