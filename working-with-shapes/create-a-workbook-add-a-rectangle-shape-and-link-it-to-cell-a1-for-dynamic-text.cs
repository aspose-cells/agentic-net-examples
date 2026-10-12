// Title: Create an Excel workbook in C# with Aspose.Cells, add a rectangle shape and bind its text to cell A1 for dynamic updates
// AI Prompts: Generate a new workbook, insert a rectangle shape at row 2 column 2, set the shape’s Text property to the formula "=A1" so it reflects the cell’s value, center the text horizontally, and save as output.xlsx. | Using Aspose.Cells for .NET, create a worksheet, add a Rectangle drawing, bind its text to cell A1 with a formula, apply center alignment, and export the file.
// Common Searches: how to bind a shape's text to a cell value using Aspose.Cells C# | Aspose.Cells add rectangle shape and link text to A1 dynamically | C# Aspose.Cells set shape text formula =A1 | center text inside a shape Aspose.Cells .NET example | link Excel shape text to worksheet cell with Aspose.Cells API
// Tags: add rectangle shape Aspose.Cells | shape text formula binding Aspose.Cells | dynamic shape text from cell C# | center shape text alignment Aspose.Cells | save workbook as XLSX Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, adds a rectangle shape at row 2/column 2, binds its Text property to cell A1 with the formula "=A1", centers the text, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape (type: Rectangle) at row 2, column 2 with no offset, height 100, width 200
        Shape rectangle = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            2,   // upper left row
            2,   // upper left column
            0,   // upper left row offset (pixels)
            0,   // upper left column offset (pixels)
            100, // height (pixels)
            200  // width (pixels)
        );

        // Link the shape's text to cell A1 using a formula; the shape will display the cell's value dynamically
        rectangle.Text = "=A1";

        // Optional: center the text inside the shape
        rectangle.TextHorizontalAlignment = TextAlignmentType.Center;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
