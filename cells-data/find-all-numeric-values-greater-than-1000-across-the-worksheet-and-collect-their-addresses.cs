// Title: Find and list cell addresses with numeric values greater than 1000 using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over all populated cells in a worksheet and returns a collection of cell names where the numeric value exceeds 1000. | Adjust the example to also capture cells whose value equals 1000 and output their references together with the >1000 cells.
// Common Searches: aspnet aspose.cells how to get addresses of cells with values over 1000 | c# iterate through worksheet used range and filter numbers greater than 1000 | retrieve list of Excel cell references where numeric value exceeds 1000 using Aspose.Cells | example code to collect high-value cell addresses from an .xlsx file in C#
// Tags: scan worksheet cells Aspose.Cells | filter numeric cells >1000 C# | collect cell references Aspose.Cells | numeric threshold extraction Excel .NET | iterate cells by value Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample loads an Excel workbook, iterates through every populated cell in the first worksheet, checks if the cell contains a numeric value greater than 1000, adds the cell's address (e.g., "B5") to a list, prints the collected addresses, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // List to store addresses of numeric values greater than 1000
        List<string> addresses = new List<string>();

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        // Iterate through all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];
                // Check if the cell contains a numeric value > 1000
                if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > 1000)
                {
                    // Add the cell address (e.g., "B5") to the list
                    addresses.Add(cell.Name);
                }
            }
        }

        // Output the collected addresses (can be replaced with other handling)
        foreach (string address in addresses)
        {
            Console.WriteLine(address);
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
