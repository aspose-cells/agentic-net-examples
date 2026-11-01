// Title: How to filter out rows with Status = "Inactive" in an Excel worksheet before exporting to JSON using Aspose.Cells for .NET
// Common Searches: Aspose.Cells remove rows where column value is Inactive before saving as JSON in C# | C# filter Excel rows by status column using Aspose.Cells and export to JSON | Delete rows with specific cell text in Aspose.Cells workbook prior to JSON conversion | How to exclude inactive records from Excel when generating JSON with Aspose.Cells | Aspose.Cells C# example for row deletion based on cell content before JSON export
// Tags: Aspose.Cells row deletion by cell value | filter Excel rows before JSON conversion Aspose.Cells | C# Aspose.Cells SaveFormat.Json with filtered worksheet | exclude inactive records Excel Aspose.Cells | reverse row iteration for safe deletion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates rows in reverse, deletes any row whose Status column contains "Inactive", and then saves the cleaned worksheet to a JSON file using Aspose.Cells SaveFormat.Json.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Column index for the "Status" column (e.g., column B => index 1)
                int statusColumnIndex = 1;

                // Determine the data range
                int minRow = sheet.Cells.MinDataRow;
                int maxRow = sheet.Cells.MaxDataRow;

                // Remove rows where Status = "Inactive"
                for (int rowIndex = maxRow; rowIndex >= minRow; rowIndex--)
                {
                    try
                    {
                        Cell statusCell = sheet.Cells[rowIndex, statusColumnIndex];
                        if (statusCell != null && statusCell.Type == CellValueType.IsString &&
                            string.Equals(statusCell.StringValue, "Inactive", StringComparison.OrdinalIgnoreCase))
                        {
                            // Delete the entire row
                            sheet.Cells.DeleteRows(rowIndex, 1);
                        }
                    }
                    catch (Exception exRow)
                    {
                        Console.WriteLine($"Warning: Could not process row {rowIndex}. {exRow.Message}");
                    }
                }

                // Save the workbook as JSON
                string outputPath = "output.json";
                workbook.Save(outputPath, SaveFormat.Json);
                Console.WriteLine($"Workbook saved as JSON to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
