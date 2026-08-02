// Title: Add a rectangle shape with text and an external hyperlink using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, accesses the first worksheet, inserts a rectangle shape at row 2/column 2, sets its displayed text, attaches a hyperlink to an external URL, and saves the file as ShapeWithHyperlink.xlsx.
// Keywords: Aspose.Cells shape hyperlink C# | add rectangle shape Excel Aspose | Excel shape with clickable link .NET | Aspose.Cells custom text shape | C# generate Excel with hyperlink shape
// Common Searches: Aspose.Cells add clickable shape with URL | C# insert rectangle shape and hyperlink in Excel | How to set a hyperlink on a shape using Aspose.Cells | Create a button shape in Excel that opens a website | Aspose.Cells shape AddHyperlink example
// Developer Intent: Insert a shape, display custom text, and bind it to an external web address in an Excel workbook.
// Use Cases: Add a call‑to‑action button in automated reports that opens documentation. | Place a branded banner in a template that redirects to a product page. | Provide one‑click access to support resources from within a generated spreadsheet.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape, set its text, attach a hyperlink, and save the workbook. | Explain how to modify the shape's size, position, and hyperlink target in the provided Aspose.Cells example. | Show how to create multiple shapes, each linking to a different URL, using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, accesses the first worksheet, inserts a rectangle shape at row 2/column 2, sets its displayed text, attaches a hyperlink to an external URL, and saves the file as ShapeWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
        var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

        // Set the custom text displayed inside the shape
        shape.Text = "Click for additional resources";

        // Attach a hyperlink that points to an external website
        shape.AddHyperlink("https://www.example.com/resources");

        // Save the workbook to a file
        workbook.Save("ShapeWithHyperlink.xlsx");
    }
}
