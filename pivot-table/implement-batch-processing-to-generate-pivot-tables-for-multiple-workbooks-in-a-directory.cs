// Title: Create pivot tables in every Excel workbook of a folder using Aspose.Cells for .NET (C# batch processing)
// AI Prompts: Generate a C# console program that enumerates all .xlsx files in a given directory, inserts a new worksheet with a pivot table based on the first two columns, refreshes the pivot, and writes the workbook to an output folder using Aspose.Cells. | Modify the batch pivot generator to accept command‑line arguments for the source and destination folders and for the indices of the row and data fields to include in each pivot table. | Add code that applies a built‑in pivot table style (e.g., PivotStyleMedium9) to each generated pivot sheet and logs any files that fail to process.
// Common Searches: aspnet add pivot table to each Excel file in a directory using Aspose.Cells | c# batch generate pivot tables for multiple workbooks Aspose.Cells example | how to refresh pivot tables after programmatic creation with Aspose.Cells | process all .xlsx files in a folder and create pivot tables automatically c# | aspose.cells create pivot table from used range in batch script
// Tags: batch create pivot tables Aspose.Cells C# | insert pivot table via Aspose.Cells | update pivot tables after creation Aspose.Cells | iterate over .xlsx files in folder C# | save modified workbooks to output folder

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, adds a new worksheet containing a simple pivot table built from the first two columns of the used range, refreshes all pivot tables, and saves the updated workbook to an output directory.
class BatchPivotGenerator
{
    static void Main()
    {
        // Directory containing the source workbooks
        string inputDir = @"C:\InputWorkbooks";

        // Directory where the processed workbooks will be saved
        string outputDir = @"C:\OutputWorkbooks";

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory does not exist: {inputDir}");
            return;
        }

        // Process each .xlsx file in the input directory
        foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            // Skip if the file is somehow missing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Use the first worksheet as the data source
                Worksheet dataSheet = workbook.Worksheets[0];

                // Determine the used range of the data sheet
                AsposeRange usedRange = dataSheet.Cells.MaxDisplayRange;
                string sourceData = $"={dataSheet.Name}!{usedRange.Address}";

                // Add a new worksheet to host the pivot table
                string pivotSheetName = "Pivot_" + Path.GetFileNameWithoutExtension(filePath);
                Worksheet pivotSheet = workbook.Worksheets.Add(pivotSheetName);

                // Add a pivot table to the new worksheet
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "PivotTable1");
                PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

                // Simple configuration: first column as row field, second column as data field
                if (usedRange.ColumnCount >= 2)
                {
                    pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Row field
                    pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Data field
                }

                // Refresh all pivot tables in the workbook to ensure they reflect the current data
                workbook.Worksheets.RefreshPivotTables();

                // Save the modified workbook to the output directory
                string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                workbook.Save(outputPath);

                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
