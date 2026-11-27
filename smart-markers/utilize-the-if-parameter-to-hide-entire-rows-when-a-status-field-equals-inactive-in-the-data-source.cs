// Title: Hide entire rows in an Excel worksheet when the status column equals 'Inactive' using Aspose.Cells for C#
// Common Searches: C# Aspose.Cells hide rows where column B value is Inactive | How to programmatically hide Excel rows based on a status field with Aspose.Cells | Using Cells.HideRow to hide rows conditionally in a .xlsx file via Aspose.Cells C# | Aspose.Cells example for hiding rows when a cell contains 'Inactive' | Smart markers If parameter to hide rows with specific status in C# Aspose.Cells
// Tags: hide rows based on cell value Aspose.Cells C# | Cells.HideRow method Excel Aspose.Cells | conditional row visibility .xlsx C# Aspose | filter rows by status column Aspose.Cells | programmatic row hiding Excel workbook Aspose.Cells | smart markers If parameter row hide example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, iterates through rows of the first worksheet, checks the status column (B) for the value "Inactive", hides matching rows using the Cells.HideRow method, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0]; // first worksheet

                // Column B (index 1) holds the status field
                const int statusColumnIndex = 1;
                const int firstDataRow = 1; // assuming row 0 contains headers
                int lastDataRow = worksheet.Cells.MaxDataRow;

                // Hide rows where the status is "Inactive"
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    Cell statusCell = worksheet.Cells[row, statusColumnIndex];
                    if (statusCell.StringValue.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                    {
                        // Use Cells.HideRow to hide the entire row
                        worksheet.Cells.HideRow(row);
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
