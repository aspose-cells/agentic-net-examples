// Title: C# – Link a shape’s text to cell K3 with automatic refresh using Aspose.Cells
// Description: Demonstrates how to add a rectangle shape to a worksheet, bind its text to the absolute cell reference K3, trigger an update so the shape reflects the current cell value, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shape linked cell | shape.LinkedCell property | auto‑refresh shape text | bind shape to Excel cell | dynamic shape label Aspose | .NET Excel automation example | update shape value from cell
// Common Searches: Aspose.Cells link shape to cell K3 C# | how to bind shape text to worksheet cell using Aspose | shape.UpdateSelectedValue after linking cell | C# example for dynamic shape text in Excel | Aspose.Cells shape linkedcell usage
// Developer Intent: The developer needs a shape whose displayed text automatically mirrors the value of cell K3 whenever that cell changes.
// Use Cases: KPI dashboards where shapes show live metric values from specific cells. | Report templates with dynamic labels inside shapes that stay in sync with worksheet data. | Excel‑based forms that use shapes as visual placeholders for cell content.
// AI Prompts: Provide C# code that links a shape’s text to a cell and keeps it updated with Aspose.Cells. | Show how to bind multiple shapes to different cells and refresh them automatically in a .NET workbook. | Explain the difference between using shape.LinkedCell and setting shape.Text directly in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a worksheet, bind its text to the absolute cell reference K3, trigger an update so the shape reflects the current cell value, and save the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), height, width
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 100);

        // Link the shape's text to cell K3 (absolute reference)
        shape.LinkedCell = "$K$3";

        // Refresh the shape so it displays the current cell value
        shape.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("LinkedShape.xlsx");
    }
}
