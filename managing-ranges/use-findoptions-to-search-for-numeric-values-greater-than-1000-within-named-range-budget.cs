// Title: Find numeric cells greater than 1000 in the "Budget" named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find cells with values above 1000 in a named range | how to iterate over a named range and filter numeric values using Aspose.Cells | search Excel workbook for numbers greater than a threshold with Aspose.Cells .NET | retrieve high budget amounts from a named range using Aspose.Cells API | C# code example for locating values over 1000 in the Budget range with Aspose.Cells
// Tags: find numeric cells in named range Aspose.Cells | filter Excel cells by value threshold C# | search named range for values greater than 1000 | Aspose.Cells numeric value comparison | retrieve high budget values from workbook

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the "Budget" named range, iterates through its cells, identifies numeric values exceeding 1000, outputs their addresses and values, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the named range "Budget"
                AsposeRange budgetRange = workbook.Worksheets.GetRangeByName("Budget");
                if (budgetRange == null)
                {
                    Console.WriteLine("Named range \"Budget\" not found in the workbook.");
                    return;
                }

                // Iterate through all cells in the defined range and find values > 1000
                int startRow = budgetRange.FirstRow;
                int endRow = startRow + budgetRange.RowCount - 1;
                int startCol = budgetRange.FirstColumn;
                int endCol = startCol + budgetRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > 1000)
                        {
                            // Example action: output the address and value
                            Console.WriteLine($"Found value {cell.DoubleValue} at {cell.Name}");
                        }
                    }
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
