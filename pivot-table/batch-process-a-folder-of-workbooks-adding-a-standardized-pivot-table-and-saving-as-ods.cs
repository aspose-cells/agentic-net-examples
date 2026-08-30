// Title: Batch add a standardized pivot table to each .xlsx workbook in a folder and export them as ODS files with Aspose.Cells for .NET
// AI Prompts: Write a C# console app that iterates over every .xlsx file in a specified directory, creates a pivot table on the first worksheet (first column as row field, second column as data field) at cell E5, refreshes all pivots, and saves each workbook as an .ods file using Aspose.Cells OdsSaveOptions with pivots preserved. | Generate Aspose.Cells code to bulk‑process Excel workbooks: load each file, detect the used range, add a pivot table named "StandardPivot", assign row and data fields, refresh the pivots, and export to ODS while keeping the pivot tables intact.
// Common Searches: c# Aspose.Cells add pivot table to multiple Excel files in a folder | how to convert a batch of .xlsx files to .ods preserving pivot tables using Aspose.Cells | automate creation of standard pivot tables in Excel workbooks with Aspose.Cells .NET | bulk process Excel workbooks and export to ODS format with pivot tables included | Aspose.Cells OdsSaveOptions keep pivot tables when saving as ODS
// Tags: batch pivot table creation Aspose.Cells | xlsx to ods conversion with Aspose.Cells | add pivot table programmatically C# | OdsSaveOptions include pivots | process multiple workbooks .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The C# program scans an input folder for .xlsx files, loads each workbook with Aspose.Cells, determines the used data range, adds a pivot table named "StandardPivot" at cell E5 on the first worksheet (first column as row field, second column as data field), refreshes all pivots, and saves the workbook as an .ods file in an output folder using OdsSaveOptions configured to retain the pivot tables.
class BatchPivotProcessor
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
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Use the first worksheet as the source data sheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range (e.g., A1:D10)
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;
            string sourceRange = $"A1:{CellIndexToName(maxRow, maxCol)}";

            // Destination cell for the new pivot table
            string destCell = "E5";

            // Add a standardized pivot table
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceRange, destCell, "StandardPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot: first column as row field, second column as data field (if present)
            if (maxCol >= 1)
            {
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            }

            // Refresh all pivot tables in the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Save as ODS, ensuring pivot tables are included
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.IgnorePivotTables = false; // include pivot tables

            string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".ods";
            string outputPath = Path.Combine(outputFolder, outputFileName);
            workbook.Save(outputPath, saveOptions);
        }
    }

    // Helper method: converts zero‑based row/column indices to an Excel cell address (e.g., 0,0 -> A1)
    static string CellIndexToName(int row, int col)
    {
        // Convert column index to letters (0 -> A, 27 -> AB, etc.)
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
