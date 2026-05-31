using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsShowFormulasDemo
{
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the worksheet by its name
                const string sheetName = "DataSheet";
                Worksheet worksheet = workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Enable formula display for the worksheet
                worksheet.ShowFormulas = true;

                // Iterate over used cells and print the formula text (or value if no formula)
                Cells cells = worksheet.Cells;
                AsposeRange usedRange = cells.MaxDisplayRange; // get the used range

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // When ShowFormulas is true, StringValue returns the formula text
                        Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
                    }
                }

                // Save the workbook with formulas displayed
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}