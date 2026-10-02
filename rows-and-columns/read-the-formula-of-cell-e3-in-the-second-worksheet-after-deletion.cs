// Title: Retrieve the formula in cell E3 of the second worksheet after removing the first worksheet with Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, delete the sheet at index 0, then extract the Formula property of cell E3 on the remaining second sheet using Aspose.Cells in C#. | Using Aspose.Cells for .NET, open a workbook, remove the first worksheet, and read the formula string stored in cell E3 of the next worksheet.
// Common Searches: Aspose.Cells C# get formula from a cell after deleting a worksheet | read formula from E3 after removing first sheet in a .NET workbook | example code to delete first worksheet and retrieve cell formula with Aspose.Cells | how to access cell formula when worksheet indices change in Aspose.Cells
// Tags: Aspose.Cells delete first sheet | Aspose.Cells read cell formula | C# worksheet index handling Aspose.Cells | cell.Formula property usage | Aspose.Cells workbook manipulation

using Aspose.Cells;
using System;

// The example loads an Excel file, removes the first worksheet, accesses the worksheet now at index 1, reads the formula from cell E3 using the Cell.Formula property, and prints the formula to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Delete the first worksheet (index 0)
        workbook.Worksheets.RemoveAt(0);

        // Access the second worksheet (now at index 1)
        Worksheet secondSheet = workbook.Worksheets[1];

        // Retrieve cell E3
        Cell targetCell = secondSheet.Cells["E3"];

        // Read the formula stored in the cell
        string formula = targetCell.Formula;

        // Display the formula
        Console.WriteLine("Formula in E3 of the second worksheet: " + formula);
    }
}
