// Title: Aspose.Cells for .NET – Access a Worksheet and Retrieve a Cell (B2) in C#
// Description: C# example that creates a new Workbook, selects the first Worksheet, gets a reference to cell B2, writes "Hello World", prints the cell address and value, and saves the file as AccessCellDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells get cell by address | C# Aspose.Cells worksheet access | retrieve cell reference Aspose.Cells | write value to cell B2 | save workbook Aspose.Cells .NET | Aspose.Cells cell object | Aspose.Cells sample code
// Common Searches: Aspose.Cells how to get a cell reference in C# | C# Aspose.Cells access first worksheet | retrieve cell B2 using Aspose.Cells | write data to a specific cell Aspose.Cells .NET | save workbook after editing a cell Aspose.Cells
// Developer Intent: The developer needs to locate a specific cell within a worksheet, modify its contents, and persist the changes in an Excel file using Aspose.Cells for .NET.
// Use Cases: Insert a header or label into cell B2 of a newly created workbook. | Read, display, and verify the value of a targeted cell after updating it. | Programmatically generate an Excel file, set values in precise cells, and export the file to disk.
// AI Prompts: Generate C# code that opens an existing workbook, selects the second worksheet, and writes "Total" to cell C5 with Aspose.Cells. | Show how to loop through a range starting at B2 and apply a numeric format to each cell using Aspose.Cells. | Explain how to retrieve a cell's style, change its font to bold and background color, and then save the workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a new Workbook, selects the first Worksheet, gets a reference to cell B2, writes "Hello World", prints the cell address and value, and saves the file as AccessCellDemo.xlsx using Aspose.Cells.
class AccessCellDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain a reference to the target cell B2
        Cell targetCell = worksheet.Cells["B2"];

        // Put a sample value into the cell
        targetCell.PutValue("Hello World");

        // Display the cell address and its value
        Console.WriteLine($"Cell {targetCell.Name} contains: {targetCell.StringValue}");

        // Save the workbook (lifecycle: save)
        workbook.Save("AccessCellDemo.xlsx");
    }
}
