// Title: How to strip all table formatting by setting TableToRangeOptions.LastRow to 0 when converting an Excel table to a range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert Excel table to range without keeping any formatting | set TableToRangeOptions.LastRow to zero to remove table styles in Aspose.Cells | remove table formatting during ListObject to range conversion using Aspose.Cells .NET | C# Aspose.Cells convert ListObject to normal range and discard style | how to disable table style preservation when converting to range with Aspose.Cells
// Tags: TableToRangeOptions LastRow zero | Aspose.Cells table-to-range conversion | remove table style during conversion C# | convert ListObject to range without preserving style | strip Excel table formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, checks for a worksheet and a table, creates a TableToRangeOptions object with LastRow set to 0 to eliminate all table formatting, converts the first ListObject to a regular range using those options, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("Error: No tables (ListObjects) found on the first worksheet.");
                return;
            }

            // Assume the first table on the worksheet is the target
            ListObject table = sheet.ListObjects[0];

            // Create conversion options; setting LastRow to 0 removes all formatting rows during conversion
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = 0
            };

            // Convert the table to a normal range using the specified options
            table.ConvertToRange(options);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
