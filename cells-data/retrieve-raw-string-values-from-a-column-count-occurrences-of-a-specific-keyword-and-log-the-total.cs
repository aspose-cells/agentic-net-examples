// Title: Count occurrences of a specific keyword in an Excel column using Aspose.Cells for .NET (C#)
// Common Searches: how to count a word in a specific Excel column with Aspose.Cells C# | retrieve raw string value from a cell in Aspose.Cells .NET | iterate over used rows in a worksheet using Aspose.Cells | convert column index to column letter Aspose.Cells C# | log keyword occurrence count to console with Aspose.Cells
// Tags: keyword occurrence counting Aspose.Cells | raw string value extraction C# Excel | column index to name conversion Aspose.Cells | used rows iteration Aspose.Cells workbook | console logging of count Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, accesses the first worksheet, iterates through all used rows of a chosen column, reads each cell's raw string value, increments a counter when the value contains a specified keyword, converts the column index to its letter name, and writes the total occurrence count to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet sheet = workbook.Worksheets[0];

            // Column index to inspect (0 = A, 1 = B, etc.)
            int columnIndex = 0; // Column A

            // Keyword to count
            string keyword = "TargetKeyword";

            // Counter for occurrences
            int occurrenceCount = 0;

            // Iterate through all used rows in the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                // Retrieve the raw string value from the cell
                string cellValue = sheet.Cells[row, columnIndex].StringValue;

                // If the cell contains the keyword, increment the counter
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(keyword))
                {
                    occurrenceCount++;
                }
            }

            // Log the total count
            string columnName = CellsHelper.ColumnIndexToName(columnIndex);
            Console.WriteLine($"Total occurrences of \"{keyword}\" in column {columnName}: {occurrenceCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
