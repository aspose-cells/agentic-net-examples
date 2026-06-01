using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace BatchPivotTableGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing source workbooks
            string inputDir = @"C:\InputWorkbooks";
            // Output directory where modified workbooks will be saved
            string outputDir = @"C:\OutputWorkbooks";

            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Process each .xlsx file in the input directory
            foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
            {
                // Verify the source file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine the used range of the worksheet
                        AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
                        if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                            continue; // Skip empty sheets

                        // Build the source data string in A1 style (e.g., =Sheet1!A1:B10)
                        string sourceData = $"={sheet.Name}!{usedRange.Address}";

                        // Destination cell for the new pivot table (start after a one‑column gap)
                        int destColumn = usedRange.ColumnCount + 2;
                        string destCell = CellsHelper.CellIndexToName(0, destColumn);

                        // Unique pivot table name
                        string pivotName = $"Pivot_{sheet.Name}";

                        // Add a new pivot table
                        int pivotIndex = sheet.PivotTables.Add(sourceData, destCell, pivotName);
                        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                        // Configure fields: first column as row field, second column as data field
                        if (usedRange.ColumnCount >= 2)
                        {
                            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // first column
                            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // second column
                        }

                        // Refresh the newly created pivot table
                        pivotTable.RefreshData();
                    }

                    // Refresh all pivot tables in the workbook (optional)
                    workbook.Worksheets.RefreshPivotTables();

                    // Save the modified workbook to the output directory
                    string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch pivot table generation completed.");
        }
    }
}