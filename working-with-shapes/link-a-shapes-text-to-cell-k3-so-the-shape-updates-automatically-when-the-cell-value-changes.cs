// Title: How to bind a rectangle shape’s text to cell K3 for automatic updates using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape text to reference a worksheet cell | link rectangle shape text to Excel cell value with Aspose.Cells .NET | auto‑update shape text when cell K3 changes using Aspose.Cells | assign formula =K3 to shape text in Aspose.Cells C# example | bind shape text to cell in generated Excel file with Aspose.Cells
// Tags: shape text formula linking Aspose.Cells | rectangle shape cell reference .NET | auto updating shape text Excel | shape linked to worksheet cell value | C# shape text =K3 example

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, writing a value to cell K3, adding a rectangle shape, assigning the shape’s Text property to the formula "=K3" so the shape displays the cell’s content and updates automatically, centering the text, and saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample value into cell K3 (column 10, row 2 zero‑based)
        sheet.Cells["K3"].PutValue("Initial Text");

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 10, 0, 0, 100, 200);

        // Link the shape's text to cell K3 so it updates automatically
        // Setting the Text property to a formula makes the shape display the cell's value
        shape.Text = "=K3";

        // Optionally format the shape (e.g., center the text)
        shape.TextHorizontalAlignment = TextAlignmentType.Center;
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook (lifecycle rule: save)
        workbook.Save("LinkedShape.xlsx");
    }
}
