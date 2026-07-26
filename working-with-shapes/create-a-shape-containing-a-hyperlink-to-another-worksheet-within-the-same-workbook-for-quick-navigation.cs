// Title: C# – Add a rectangle shape with an internal worksheet hyperlink using Aspose.Cells
// Description: Demonstrates how to create a workbook, rename the first sheet, add a second sheet, insert a rectangle shape on the first sheet, and attach an internal hyperlink that jumps to cell A1 of the second sheet. The shape includes a screen tip and the file is saved as an XLSX document.
// Keywords: Aspose.Cells shape hyperlink | internal worksheet link C# | add rectangle shape Aspose.Cells | navigate between sheets shape | Aspose.Cells C# example
// Common Searches: Aspose.Cells add hyperlink to shape | C# shape linking to another worksheet | internal hyperlink for rectangle shape Aspose.Cells | how to create navigation button in Excel with Aspose.Cells
// Developer Intent: Insert a shape on a worksheet and bind it to an internal hyperlink that opens a specific cell on a different sheet.
// Use Cases: Create a dashboard button that opens a detailed report sheet. | Build a table‑of‑contents page with clickable shapes for each section. | Design an interactive workbook where shapes act as menu items to switch worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape on Sheet1 linking to cell B5 on Sheet3 and includes a custom screen tip. | Write a method that loops through a list of sheet names, creates rectangle shapes on a "Menu" sheet, and assigns each shape an internal hyperlink to the corresponding sheet's A1 cell.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, rename the first sheet, add a second sheet, insert a rectangle shape on the first sheet, and attach an internal hyperlink that jumps to cell A1 of the second sheet. The shape includes a screen tip and the file is saved as an XLSX document.
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
        // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, width, height
        Shape shape = homeSheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);
        shape.Text = "Go to Details";

        // Add a hyperlink to the shape that points to cell A1 of the Details sheet
        // Internal hyperlink format: 'SheetName'!CellReference
        Hyperlink hyperlink = shape.AddHyperlink("'Details'!A1");
        hyperlink.ScreenTip = "Click to navigate to Details sheet";

        // Save the workbook
        workbook.Save("ShapeWithInternalHyperlink.xlsx");
    }
}
