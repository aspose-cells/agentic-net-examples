using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a date value in cell A1
        worksheet.Cells["A1"].PutValue(new DateTime(2023, 12, 31));

        // Use TEXT function with a custom date format in cell B1
        worksheet.Cells["B1"].Formula = @"=TEXT(A1,""dd-mmm-yyyy"")";

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
        Shape rect = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Link the shape to cell B1 using SetLinkedCell method
        rect.SetLinkedCell("$B$1", false, false);

        // Verify the linked cell address
        Console.WriteLine("Shape's LinkedCell: " + rect.LinkedCell);

        // Verify the value retrieved from the linked cell
        string linkedValue = worksheet.Cells["B1"].StringValue;
        Console.WriteLine("Linked cell value (TEXT formatted date): " + linkedValue);

        // Save the workbook (optional verification)
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}