// Title: Auto‑fit a worksheet row after updating a cell value with Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet load workbook modify cell and auto fit row height | how to use worksheet.autofitrow in Aspose.Cells after changing cell text | adjust Excel row height automatically with Aspose.Cells C# example | save updated Excel file with auto‑fitted rows using Aspose.Cells | auto fit row after updating cell value Aspose.Cells .NET tutorial
// Tags: Aspose.Cells C# AutoFitRow method | modify Excel cell then scale row height | worksheet row height auto‑adjust .NET | save workbook after row auto‑adjust Aspose | programmatic row height scaling Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook, changes the value of cell B2 on the first worksheet, automatically adjusts the height of the row containing that cell with AutoFitRow, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify the value of cell B2 (row 1, column 1)
        Cell cell = worksheet.Cells["B2"];
        cell.PutValue("New long text that may require row height adjustment");

        // Auto‑fit the row that contains the modified cell
        worksheet.AutoFitRow(cell.Row);

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
