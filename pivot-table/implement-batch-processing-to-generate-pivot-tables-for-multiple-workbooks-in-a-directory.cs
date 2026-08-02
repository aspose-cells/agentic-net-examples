// Title: Batch generate pivot tables for multiple Excel workbooks using Aspose.Cells (.NET)
// Description: C# program that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, creates a pivot table at E1 using the first two columns (or a single column when only one exists), refreshes all pivots, and saves the updated files to an output directory with error handling.
// Keywords: Aspose.Cells pivot table C# | batch create pivot tables | process multiple Excel files .NET | automate pivot table generation | refresh pivot tables Aspose | iterate Excel files directory | C# Excel automation
// Common Searches: how to add a pivot table to every workbook in a folder using Aspose.Cells | C# batch generate pivot tables for .xlsx files | Aspose.Cells create pivot table from used range | refresh all pivot tables programmatically Aspose | automate pivot table insertion across multiple Excel files
// Developer Intent: Automatically add a standard pivot table to each Excel workbook in a specified directory and save the modified files.
// Use Cases: Generate a uniform sales‑summary pivot in all regional workbooks before consolidating monthly reports. | Prepare data‑analysis templates by inserting a predefined pivot layout into every uploaded Excel file in a batch workflow. | Ensure financial models contain up‑to‑date pivot calculations by refreshing and persisting pivots across a collection of workbooks.
// AI Prompts: Write C# code with Aspose.Cells that adds a pivot table to all .xlsx files in a folder, using the first two columns as row and data fields, and saves the results to an output directory. | Extend the batch pivot generator to log each processed file, skip non‑Excel files, and produce a summary of successes and failures. | Show how to set the pivot table name and placement dynamically based on each workbook’s filename when processing a directory with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace BatchPivotGenerator
{
    // C# program that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, creates a pivot table at E1 using the first two columns (or a single column when only one exists), refreshes all pivots, and saves the updated files to an output directory with error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing source workbooks
            string inputDirectory = @"C:\InputWorkbooks";
            // Directory where processed workbooks will be saved
            string outputDirectory = @"C:\OutputWorkbooks";

            // Ensure input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Process each .xlsx file in the input directory
            foreach (string filePath in Directory.GetFiles(inputDirectory, "*.xlsx"))
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Work with the first worksheet (assumed to contain the source data)
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Determine the used range of the worksheet to use as the pivot source
                    Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

                    // Build the source data reference in A1 style (e.g., =Sheet1!A1:C10)
                    string sourceData = $"={worksheet.Name}!{usedRange.Address}";

                    // Add a new pivot table at cell E1 with a default name
                    int pivotIndex = worksheet.PivotTables.Add(sourceData, "E1", "PivotTable1");
                    PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                    // Configure the pivot table:
                    // - First column as Row field
                    // - Second column as Data field (if present)
                    if (usedRange.ColumnCount >= 2)
                    {
                        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // First column
                        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Second column
                    }
                    else if (usedRange.ColumnCount == 1)
                    {
                        // If only one column exists, use it as both Row and Data
                        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                        pivotTable.AddFieldToArea(PivotFieldType.Data, 0);
                    }

                    // Refresh all pivot tables in the workbook
                    workbook.Worksheets.RefreshPivotTables();

                    // Save the modified workbook to the output directory
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
}
