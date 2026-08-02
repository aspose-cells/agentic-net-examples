// Title: C# – Concatenate raw string values from a cell range and write to a summary cell using Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, fills the range A1:C3 with mixed‑type data, reads each cell's raw StringValue, concatenates the values, writes the combined text to cell E1, and saves the file as ConcatenatedRangeStrings.xlsx. It demonstrates range creation, cell iteration, and writing a summary string.
// Keywords: Aspose.Cells C# example | concatenate cell range | raw string values Aspose.Cells | StringValue Excel .NET | write summary cell | CreateRange Aspose.Cells | iterate over cells | mixed data types Excel | Excel automation .NET | concatenation performance StringBuilder
// Common Searches: how to concatenate values from a range with Aspose.Cells .NET | read raw string from mixed type cells Aspose.Cells | write concatenated text to another cell using Aspose.Cells | Aspose.Cells C# loop through range cells | concatenate Excel cells programmatically C#
// Developer Intent: Read the raw string representation of every cell in a specified range, combine them into a single text string, and store the result in a designated summary cell.
// Use Cases: Generate a unique key by merging header cells for lookup tables. | Create a single log entry from a block of mixed‑type data before exporting. | Display a compact summary of a table section on a dashboard cell.
// AI Prompts: Show a C# Aspose.Cells snippet that concatenates the StringValue of all cells in a given range and writes the result to a target cell. | Explain how to skip empty cells or insert a delimiter while concatenating range values with Aspose.Cells. | Demonstrate how to improve performance for large ranges by using StringBuilder in the concatenation loop.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, fills the range A1:C3 with mixed‑type data, reads each cell's raw StringValue, concatenates the values, writes the combined text to cell E1, and saves the file as ConcatenatedRangeStrings.xlsx. It demonstrates range creation, cell iteration, and writing a summary string.
    public class ConcatenateRangeStrings
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate a sample range (A1:C3) with various values
                cells["A1"].PutValue("Hello");
                cells["B1"].PutValue(123);               // numeric value
                cells["C1"].PutValue(DateTime.Now);      // date/time value
                cells["A2"].PutValue("World");
                cells["B2"].PutValue(true);              // boolean value
                cells["C2"].PutValue("Aspose");
                cells["A3"].PutValue("Cells");
                cells["B3"].PutValue(45.67);             // double value
                cells["C3"].PutValue("Demo");

                // Define the range to read (A1:C3)
                int firstRow = 0;      // zero‑based index for row 1
                int firstColumn = 0;   // zero‑based index for column A
                int totalRows = 3;
                int totalColumns = 3;
                AsposeRange range = cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

                // Concatenate the raw string values of each cell in the range
                string concatenated = string.Empty;
                for (int row = 0; row < range.RowCount; row++)
                {
                    for (int col = 0; col < range.ColumnCount; col++)
                    {
                        Cell cell = range[row, col];
                        concatenated += cell.StringValue;
                    }
                }

                // Write the concatenated result to a summary cell (e.g., E1)
                cells["E1"].PutValue(concatenated);

                // Save the workbook to a file
                workbook.Save("ConcatenatedRangeStrings.xlsx");
                Console.WriteLine("Workbook saved as ConcatenatedRangeStrings.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during execution: {ex.Message}");
            }
        }
    }
}
