// Title: How to count cells that belong to array formulas in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count cells in array formulas in an existing workbook | C# iterate worksheet used range to find array formula cells with Aspose.Cells | Get total number of array formula cells from an XLSX file using Aspose.Cells .NET | Detect and tally array formula cells in a spreadsheet with Aspose.Cells API | Count array formula cells in the first worksheet of an Excel file using C#
// Tags: count array formula cells Aspose.Cells .NET | IsArrayFormula property usage | iterate used range cells Aspose.Cells | detect array formulas in Excel workbook | worksheet cell enumeration Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates over its used range, checks each cell's IsArrayFormula flag, increments a counter for cells that are part of an array formula, and prints the total number of such cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Counter for array formulas (counts each cell that belongs to an array formula)
            int arrayFormulaCellCount = 0;

            // Get the used range of the worksheet
            Cells cells = worksheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // If the cell is part of an array formula, increment the counter
                    if (cell.IsArrayFormula)
                    {
                        arrayFormulaCellCount++;
                    }
                }
            }

            // Report the total number of cells that belong to array formulas
            Console.WriteLine($"Total cells in array formulas: {arrayFormulaCellCount}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
