// Title: Add a rectangle shape with an internal worksheet hyperlink using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, rename the first sheet, add a second sheet, place a rectangle shape on the first sheet, set its caption, attach an internal hyperlink that jumps to cell A1 of the second sheet, and save the file as ShapeWithInternalHyperlink.xlsx.
// Keywords: Aspose.Cells shape hyperlink | internal worksheet link C# | add rectangle shape Aspose.Cells | navigate between sheets programmatically | Aspose.Cells hyperlink to cell
// Common Searches: Aspose.Cells add hyperlink to shape | C# create shape that links to another worksheet | internal sheet hyperlink using Aspose.Cells | shape navigation button Aspose.Cells .NET | how to link a shape to a different sheet in Excel via code
// Developer Intent: Insert a shape into a worksheet and bind it to an internal hyperlink that opens a specific cell on another worksheet.
// Use Cases: Design a dashboard where rectangle shapes act as buttons to open detailed report sheets. | Build a table‑of‑contents page with shapes that jump to individual sections in the workbook. | Create a navigation pane that lets users switch between modules with a single click.
// AI Prompts: Generate C# code with Aspose.Cells to add a circular shape on Sheet1 that links to cell B5 on Sheet3. | Show how to add multiple shapes, each pointing to a different worksheet, using Aspose.Cells for .NET. | Explain how to modify the hyperlink target of an existing shape to reference another sheet and cell.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, rename the first sheet, add a second sheet, place a rectangle shape on the first sheet, set its caption, attach an internal hyperlink that jumps to cell A1 of the second sheet, and save the file as ShapeWithInternalHyperlink.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and rename it
        Worksheet homeSheet = workbook.Worksheets[0];
        homeSheet.Name = "Home";

        // Add a second worksheet to navigate to
        Worksheet detailsSheet = workbook.Worksheets.Add("Details");

        // Add a rectangle shape on the Home sheet
        // Parameters: upper left row, upper left column, upper left pixel offset (X), upper left pixel offset (Y), width, height
        Shape shape = homeSheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);
        shape.Text = "Go to Details";

        // Add a hyperlink to the shape that points to cell A1 of the Details sheet
        // Internal hyperlink format: 'SheetName'!CellReference
        shape.AddHyperlink("'Details'!A1");

        // Save the workbook
        workbook.Save("ShapeWithInternalHyperlink.xlsx");
    }
}
