// Title: Batch Create Pivot Tables in Multiple Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, determines the used range, adds a simple pivot table (first column as rows, second column as values), refreshes the pivot, and saves the updated file to a target directory. Includes basic error handling and folder creation.
// Keywords: Aspose.Cells batch pivot table | C# generate pivot tables programmatically | process multiple Excel files .NET | automate pivot creation Aspose | refresh pivot tables C# | folder based Excel automation | add pivot table to each workbook
// Common Searches: how to add a pivot table to every Excel file in a folder using Aspose.Cells | batch generate pivot tables C# | process all .xlsx files in a directory with Aspose.Cells | automate pivot table creation for multiple workbooks .NET | refresh pivot tables after adding them programmatically
// Developer Intent: Automatically insert a standard pivot table into each workbook within a specified directory and save the modified files.
// Use Cases: Generate a consistent sales‑summary pivot in all regional workbooks before monthly distribution. | Prepare analysis templates by adding a predefined pivot to a batch of blank Excel files stored on a shared drive. | Refresh and persist pivot tables in bulk so downstream reporting tools receive ready‑to‑use files without manual steps.
// AI Prompts: Create a C# method that iterates over all .xlsx files in a folder, adds a pivot table with Aspose.Cells, and writes the results to an output directory. | Extend the sample to accept custom row‑field and data‑field column indexes as parameters for flexible pivot layouts. | Add robust logging that records successful and failed file operations to a CSV log while continuing batch processing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace BatchPivotProcessing
{
    // Alias to avoid ambiguity with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, determines the used range, adds a simple pivot table (first column as rows, second column as values), refreshes the pivot, and saves the updated file to a target directory. Includes basic error handling and folder creation.
    public class BatchPivotProcessor
    {
        /// <param name="inputDirectory">Folder containing source workbooks.</param>
        /// <param name="outputDirectory">Folder where processed workbooks will be saved.</param>
        public void ProcessDirectory(string inputDirectory, string outputDirectory)
        {
            // Ensure output folder exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Get all .xlsx files in the input folder
            string[] files = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle rule)
                    Workbook workbook = new Workbook(filePath);

                    // Assume the first worksheet contains the source data
                    Worksheet sourceSheet = workbook.Worksheets[0];

                    // Determine the used range of the worksheet
                    AsposeRange usedRange = sourceSheet.Cells.MaxDisplayRange;
                    if (usedRange == null)
                    {
                        Console.WriteLine($"No data found in workbook: {filePath}");
                        continue;
                    }

                    int startRow = usedRange.FirstRow;
                    int startCol = usedRange.FirstColumn;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Build the source data address in A1 style (e.g., "A1:B10")
                    string sourceAddress = sourceSheet.Cells[startRow, startCol].Name + ":" +
                                           sourceSheet.Cells[endRow, endCol].Name;

                    // Destination cell for the pivot table – place it a few rows below the data
                    int destRow = endRow + 3;
                    string destCellName = $"A{destRow + 1}";

                    // Add a new pivot table to the same worksheet
                    int pivotIndex = sourceSheet.PivotTables.Add(sourceAddress, destCellName,
                        "PivotTable_" + Path.GetFileNameWithoutExtension(filePath));

                    PivotTable pivotTable = sourceSheet.PivotTables[pivotIndex];

                    // Simple configuration: first column as row field, second column as data field
                    if (pivotTable.RowFields.Count == 0 && pivotTable.DataFields.Count == 0)
                    {
                        // Add first column (index 0) to Row area
                        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                        // Add second column (index 1) to Data area, if it exists
                        if (endCol - startCol >= 1)
                            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
                    }

                    // Refresh all pivot tables in the workbook (lifecycle rule)
                    workbook.Worksheets.RefreshPivotTables();

                    // Save the modified workbook to the output folder (lifecycle rule)
                    string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputDir = @"C:\InputWorkbooks";
            string outputDir = @"C:\ProcessedWorkbooks";

            BatchPivotProcessor processor = new BatchPivotProcessor();
            processor.ProcessDirectory(inputDir, outputDir);

            Console.WriteLine("Batch pivot table processing completed.");
        }
    }
}
