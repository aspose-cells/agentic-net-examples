// Title: C# – Batch Add a Standard Pivot Table to Multiple Excel Workbooks and Export as ODS using Aspose.Cells
// Description: A complete C# example that scans a folder, loads each Excel workbook (XLSX, XLS, etc.) with Aspose.Cells, inserts a predefined pivot table on the first worksheet, refreshes all pivots, and saves the result as an ODS file while preserving the pivot. Includes folder handling, OdsSaveOptions configuration, and console feedback.
// Keywords: Aspose.Cells batch pivot table | C# add pivot to multiple workbooks | convert Excel to ODS with pivots | Aspose.Cells OdsSaveOptions | automate pivot creation C# | folder processing Aspose.Cells | GitHub Aspose.Cells example | Excel to LibreOffice ODS conversion
// Common Searches: How to add the same pivot table to all Excel files in a directory using Aspose.Cells | Batch convert Excel workbooks to ODS while keeping pivot tables in C# | Aspose.Cells programmatically create pivot table and export to ODS | C# loop through folder and add pivot table to each workbook | Aspose.Cells example for bulk ODS export with pivots
// Developer Intent: Insert an identical pivot table into every workbook in a specified folder and save each file as an ODS document.
// Use Cases: Standardize monthly sales reports by automatically adding a Row‑Data pivot before distributing ODS files to stakeholders. | Migrate a library of legacy Excel dashboards to LibreOffice‑compatible ODS format while retaining pivot functionality. | Create a batch conversion pipeline for a data‑analytics team that needs consistent pivot layouts across dozens of spreadsheets.
// AI Prompts: Generate C# code that adds a pivot table to every worksheet in each workbook and saves the file as ODS with Aspose.Cells, including robust error handling for missing ranges. | Explain how to detect the used data range dynamically before creating the pivot table in a batch process. | Show how to log each processed file, capture exceptions, and produce a summary report after batch conversion to ODS.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

// A complete C# example that scans a folder, loads each Excel workbook (XLSX, XLS, etc.) with Aspose.Cells, inserts a predefined pivot table on the first worksheet, refreshes all pivots, and saves the result as an ODS file while preserving the pivot. Includes folder handling, OdsSaveOptions configuration, and console feedback.
class BatchPivotToOds
{
    static void Main()
    {
        // Folder containing source workbooks (XLSX, XLS, etc.)
        string sourceFolder = @"C:\InputWorkbooks";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputOds";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each workbook file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder))
        {
            // Load the workbook (any supported Excel format)
            Workbook workbook = new Workbook(filePath);

            // Use the first worksheet for the pivot table
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source data range for the pivot table.
            // Here we assume data starts at A1 and occupies columns A and B.
            // Adjust the range as needed for your actual data.
            string sourceRange = "A1:B10";

            // Destination cell for the pivot table
            string destCell = "E1";

            // Add a new pivot table with a standard name
            int pivotIdx = sheet.PivotTables.Add(sourceRange, destCell, "StandardPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure the pivot: first column as Row field, second column as Data field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column A
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column B

            // Refresh all pivot tables and charts in the workbook
            workbook.Worksheets.RefreshAll();

            // Prepare ODS save options – include pivot tables in the output
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            odsOptions.IgnorePivotTables = false;

            // Build the output file path with .ods extension
            string outputPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(filePath) + ".ods");

            // Save the workbook as ODS using the specified options
            workbook.Save(outputPath, odsOptions);
        }

        Console.WriteLine("Batch processing completed.");
    }
}
