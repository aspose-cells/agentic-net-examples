using System;
using System.Collections;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class FindMaxNumericCell
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Variables to keep track of the maximum numeric value and its cell
            double maxValue = double.MinValue;
            Cell maxCell = null;

            // Get the cells enumerator and iterate through all instantiated cells
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Consider only numeric cells (int, double, DateTime)
                if (cell.IsNumericValue && cell.Value != null)
                {
                    // Convert the cell value to double for comparison
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
                Console.WriteLine($"Maximum numeric value: {maxValue} found at cell {maxCell.Name}");

                // Optionally highlight the cell with a style
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                maxCell.SetStyle(style);
            }
            else
            {
                Console.WriteLine("No numeric cells were found in the worksheet.");
            }

            // Save the workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}