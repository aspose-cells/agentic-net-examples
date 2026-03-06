using System;
using Aspose.Cells;

class GetCellByNameDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Get a cell object by its name using the string indexer of Cells
        // Example: retrieve cell C5
        Cell cell = worksheet.Cells["C5"];

        // Read and display the current value of the cell
        Console.WriteLine($"Cell {cell.Name} contains: {cell.Value}");

        // Optionally modify the cell's value
        cell.PutValue("Updated");

        // Save the workbook to a new file (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}