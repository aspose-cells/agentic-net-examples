// Title: Find the cell with the highest numeric value in an Excel worksheet using Aspose.Cells Cells enumerator in C#
// Common Searches: C# Aspose.Cells iterate all cells to get maximum numeric value | How to obtain the address of the largest number in an Excel sheet with Aspose.Cells | Find maximum numeric cell in first worksheet using Cells collection in C# | Enumerate worksheet cells and compare numeric values with Aspose.Cells | Retrieve row and column of highest numeric value in Excel using Aspose.Cells C#
// Tags: find maximum numeric value cells enumerator Aspose.Cells | retrieve cell address of max numeric value C# | enumerate worksheet cells Aspose.Cells C# | filter numeric cells Aspose.Cells | output row column of highest number Excel C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, uses the Cells enumerator to walk through every cell in the first worksheet, filters for numeric cells, tracks the greatest numeric value and its corresponding cell, and then prints the maximum value together with the cell's name, row, and column.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                double maxValue = double.MinValue;
                Cell maxCell = null;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Consider only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double currentValue = cell.DoubleValue;
                        if (currentValue > maxValue)
                        {
                            maxValue = currentValue;
                            maxCell = cell;
                        }
                    }
                }

                // Output the result
                if (maxCell != null)
                {
                    Console.WriteLine($"Maximum numeric value: {maxValue}");
                    Console.WriteLine($"Location: {maxCell.Name} (Row {maxCell.Row}, Column {maxCell.Column})");
                }
                else
                {
                    Console.WriteLine("No numeric values found in the sheet.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
