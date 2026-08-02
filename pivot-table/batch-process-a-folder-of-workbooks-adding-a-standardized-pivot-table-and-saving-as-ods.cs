// Title: Batch add a standard pivot table to Excel files and save as ODS with Aspose.Cells (C#)
// Description: Iterates through all *.xlsx files in a given folder, loads each workbook with Aspose.Cells, determines the used range on the first worksheet, inserts a pivot table named "StandardPivot" at cell F1 (first column as row field, second column as data field), refreshes the pivot, and saves the result as an ODS file preserving the pivot using OdsSaveOptions.
// Keywords: Aspose.Cells C# pivot table batch | add pivot to multiple workbooks | convert Excel to ODS with pivots | OdsSaveOptions PreservePivotTables | automate pivot creation Aspose | C# folder processing Excel files | standardized pivot layout | LibreOffice ODS export Aspose
// Common Searches: how to add a pivot table to many Excel files using Aspose.Cells | batch convert .xlsx to .ods while keeping pivots | C# code for creating pivot tables programmatically | Aspose.Cells ODS export with pivot tables | automate pivot table insertion across workbooks
// Developer Intent: Automatically insert the same pivot table into every Excel workbook in a directory and export each file to ODS format without losing the pivot data.
// Use Cases: Standardize monthly sales reports by adding a predefined pivot to each department's spreadsheet before sharing with LibreOffice users. | Migrate a legacy archive of Excel workbooks to ODS while retaining analytical pivot tables for downstream analysis. | Run a nightly job that enriches dozens of workbooks with a uniform pivot layout and stores the results as ODS for compliance archiving.
// AI Prompts: Write C# code that uses Aspose.Cells to add a pivot table to every workbook in a folder and save each as ODS with pivots preserved. | Explain the role of OdsSaveOptions.IgnorePivotTables and how to configure it to keep pivot tables during Excel‑to‑ODS conversion. | Provide error‑handling patterns for empty source ranges, missing worksheets, or file‑access issues in a batch pivot‑creation script.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Iterates through all *.xlsx files in a given folder, loads each workbook with Aspose.Cells, determines the used range on the first worksheet, inserts a pivot table named "StandardPivot" at cell F1 (first column as row field, second column as data field), refreshes the pivot, and saves the result as an ODS file preserving the pivot using OdsSaveOptions.
class BatchPivotToOds
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputOds";

        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(filePath);

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range to use as the pivot source
            int maxRow = worksheet.Cells.MaxDataRow;      // zero‑based
            int maxCol = worksheet.Cells.MaxDataColumn;   // zero‑based
            string sourceRange = $"A1:{CellIndexToName(maxRow, maxCol)}";

            // Destination cell for the new pivot table
            string destCell = "F1";

            // Add a standardized pivot table (PivotTableCollection.Add)
            int pivotIndex = worksheet.PivotTables.Add(sourceRange, destCell, "StandardPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot: first column as row field, second column as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Refresh the pivot table in the worksheet (Worksheet.RefreshPivotTables)
            worksheet.RefreshPivotTables();

            // Prepare ODS save options to include pivot tables (OdsSaveOptions.IgnorePivotTables)
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.IgnorePivotTables = false; // ensure pivot tables are saved

            // Build the output file path with .ods extension
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string outputPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

            // Save the workbook as ODS using the save options (Workbook.Save)
            workbook.Save(outputPath, saveOptions);
        }
    }

    // Helper: converts zero‑based row/column indices to an Excel cell address (e.g., 0,0 -> A1)
    static string CellIndexToName(int row, int col)
    {
        // Convert column index to letters
        string colName = "";
        int dividend = col + 1;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar('A' + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        // Row index is zero‑based, so add 1
        return $"{colName}{row + 1}";
    }
}
