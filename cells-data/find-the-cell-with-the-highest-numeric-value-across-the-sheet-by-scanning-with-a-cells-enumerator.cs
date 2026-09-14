// Title: Enumerate worksheet cells with Aspose.Cells in C# to locate the cell containing the highest numeric value
// AI Prompts: Write C# code that uses Aspose.Cells Cells.GetEnumerator() to iterate all populated cells in a worksheet and returns the address of the cell with the greatest numeric value. | Adjust the enumeration loop to treat DateTime cells as OADate numbers when comparing values for the maximum.
// Common Searches: Aspose.Cells C# find cell with maximum number in a worksheet | how to iterate over all cells in Aspose.Cells and get the highest numeric entry | C# enumerate Excel cells using Aspose.Cells to determine the largest value | retrieve address of the largest numeric cell with Aspose.Cells GetEnumerator
// Tags: cells enumerator max numeric value Aspose.Cells | find highest numeric cell C# Aspose.Cells | enumerate worksheet cells Aspose.Cells C# | compare numeric cell values Aspose.Cells | retrieve cell address max value Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

// The example loads an Excel workbook, uses Cells.GetEnumerator() to walk through every instantiated cell in the first worksheet, checks each cell for a numeric value (including DateTime as OADate), tracks the largest numeric value and its cell, prints the maximum value with the cell address, and saves the workbook unchanged.
class FindMaxNumericCell
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
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

            // Consider only cells that contain a numeric value
            if (cell.IsNumericValue && cell.Value != null)
            {
                double currentValue = cell.DoubleValue; // Works for int, double, DateTime (as OADate)

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
        }
        else
        {
            Console.WriteLine("No numeric cells were found in the worksheet.");
        }

        // Save the workbook (unchanged) to a new file
        workbook.Save("output.xlsx");
    }
}
