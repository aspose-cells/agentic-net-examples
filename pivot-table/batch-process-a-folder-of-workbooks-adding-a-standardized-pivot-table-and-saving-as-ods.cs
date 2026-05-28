using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class BatchPivotToOds
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputFolder";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputFolder";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Work with the first worksheet (customize as needed)
                Worksheet ws = workbook.Worksheets[0];

                // Determine the used range to serve as the pivot source
                int lastRow = ws.Cells.MaxDataRow;      // zero‑based index of last row with data
                int lastCol = ws.Cells.MaxDataColumn;   // zero‑based index of last column with data
                string sourceRange = $"A1:{CellIndexToName(lastRow, lastCol)}";

                // Destination cell for the new pivot table
                string destCell = "F1";

                // Add a standardized pivot table
                int pivotIndex = ws.PivotTables.Add(sourceRange, destCell, "StandardPivot");
                PivotTable pivot = ws.PivotTables[pivotIndex];

                // Configure the pivot: first column as row field, second column as data field
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Refresh the pivot table to ensure it reflects the source data
                ws.RefreshPivotTables();

                // Prepare ODS save options to include pivot tables
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    IgnorePivotTables = false // include pivot tables in the ODS file
                };

                // Build the output file path with .ods extension
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string outPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

                // Save the workbook as ODS
                workbook.Save(outPath, saveOptions);
                Console.WriteLine($"Converted: {filePath} -> {outPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }

    // Helper method: converts zero‑based row/column indices to an Excel cell address (e.g., 0,0 -> A1)
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