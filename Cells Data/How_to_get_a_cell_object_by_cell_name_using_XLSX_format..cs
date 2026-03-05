using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get a cell object by its name (e.g., "B2")
        Cell cell = worksheet.Cells["B2"];

        // Output the cell name and its current value
        Console.WriteLine($"Cell {cell.Name} contains: {cell.Value}");

        // Example: modify the cell value
        cell.PutValue("New Value");

        // Save the workbook to a new file
        workbook.Save("output.xlsx");
    }
}