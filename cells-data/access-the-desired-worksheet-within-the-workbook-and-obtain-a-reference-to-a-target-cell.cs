// Title: Load an Excel workbook, select a worksheet by name, and retrieve cell B2 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to open a workbook file, choose the worksheet named "Sheet1", and obtain a Cell object for address B2. | Show how to read and print the value of cell B2 after accessing a specific worksheet with Aspose.Cells in a .NET console application. | Provide an example of accessing a worksheet by its name and then referencing a cell by its address using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# example to read cell B2 from a named worksheet | How to open an Excel file and get a specific cell using Aspose.Cells .NET | C# code for selecting worksheet by name and retrieving cell value with Aspose.Cells | Retrieve a cell reference after loading a workbook in Aspose.Cells for .NET | Read Excel cell value from Sheet1 using Aspose.Cells library in C#
// Tags: load workbook Aspose.Cells C# | select worksheet by name Aspose.Cells | reference cell by address Aspose.Cells | extract cell value Aspose.Cells .NET | Aspose.Cells B2 cell access example

using System;
using Aspose.Cells;

// The snippet demonstrates how to load an existing Excel file (input.xlsx) with Aspose.Cells, access the worksheet named "Sheet1", obtain a reference to cell B2, and output its value to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the desired worksheet by name (or use index: workbook.Worksheets[0])
        Worksheet worksheet = workbook.Worksheets["Sheet1"];

        // Obtain a reference to the target cell (e.g., cell B2)
        Cell targetCell = worksheet.Cells["B2"];

        // Example: read the cell's value
        Console.WriteLine($"Value in {targetCell.Name}: {targetCell.Value}");
    }
}
