// Title: Create a rectangle shape linked to a cell with a custom array constant formula using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet and assign its LinkedCell property to a cell that holds an array constant formula. | Write an array constant (e.g., ={"One","Two","Three"}) into a cell, link a shape to that cell, and ensure the shape shows the first element when opened in Excel. | Generate an Excel file where a shape's displayed text reflects the value of a linked cell containing a custom array constant, then save the workbook.
// Common Searches: Aspose.Cells how to link a shape to a cell that contains an array constant | C# add rectangle shape and bind it to a cell with ={\"One\",\"Two\",\"Three\"} formula | display first element of an array constant in a shape using Aspose.Cells for .NET | set LinkedCell property for a shape to a cell with a custom array constant via code
// Tags: rectangle shape linkedcell Aspose.Cells | linkedcell array constant Aspose.Cells | shape text from linked cell Aspose.Cells | create workbook with linked shape Aspose.Cells | Aspose.Cells shape linked to cell formula

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes an array constant formula to cell A1, adds a rectangle shape at row 5 column 5, links the shape to A1 so the shape displays the first array element, optionally sets shape text, and saves the file as ShapeLinkedToArrayConstant.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Set a cell with a custom array constant formula.
            // The shape will display the first element of the array.
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "={\"One\",\"Two\",\"Three\"}";

            // Add a rectangle shape to the worksheet.
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                0,   // top offset (in pixels)
                0,   // left offset (in pixels)
                100, // height (in pixels)
                50   // width (in pixels)
            );

            // Link the shape to the cell containing the array constant.
            shape.LinkedCell = "A1";

            // Optional: set shape text (the linked cell value will appear when opened in Excel).
            shape.Text = "Linked to A1";

            // Save the workbook.
            string outputPath = "ShapeLinkedToArrayConstant.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
