// Title: Calculate the sum of numeric values in the first column of an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells, iterates through each row of the first worksheet, and returns the sum of all numeric entries in column A. | Create a reusable method in C# that accepts a workbook path and uses Aspose.Cells to compute the total of numeric cells in the first column, ignoring text or empty cells. | Adjust the example to log each numeric value found in column 0 before adding it to the cumulative sum with Aspose.Cells.
// Common Searches: Aspose.Cells C# sum values in column A of an Excel file | How to iterate rows and add numeric cells in first column with Aspose.Cells | C# calculate total of first column numbers using Aspose.Cells library | Skip non‑numeric cells when summing a column in Aspose.Cells for .NET | Retrieve and aggregate numeric data from the first column of a worksheet with Aspose.Cells
// Tags: sum first column Aspose.Cells | enumerate worksheet rows C# | filter numeric cells Aspose.Cells | calculate column total Excel .NET | iterate rows Aspose.Cells workbook

using System;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates through each row, checks if the cell in the first column contains a numeric value, adds those values to a running total, and prints the sum.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access the cells collection
        Cells cells = sheet.Cells;

        double sum = 0;

        // Enumerate each row that contains data
        foreach (Row row in cells.Rows)
        {
            // Access the cell in the first column (index 0)
            Cell cell = row[0];

            // Add to sum only if the cell holds a numeric value
            if (cell.Type == CellValueType.IsNumeric)
            {
                sum += cell.DoubleValue;
            }
        }

        Console.WriteLine("Sum of values in the first column: " + sum);
    }
}
