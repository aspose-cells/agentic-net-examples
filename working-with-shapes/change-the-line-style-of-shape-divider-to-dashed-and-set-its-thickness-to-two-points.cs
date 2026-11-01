// Title: Apply a dashed line style with 2‑point thickness to the 'Divider' shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: In C#, use Aspose.Cells to locate the shape named 'Divider' on the first worksheet and set its Line.DashStyle to Dash with a Line.Weight of 2 points. | Programmatically give a specific worksheet shape a 2‑point dashed border by configuring the Shape.Line properties through Aspose.Cells.
// Common Searches: Aspose.Cells C# set dashed line style for a worksheet shape | How to change shape line weight to 2 points using Aspose.Cells | Modify border dash pattern of a named shape in Excel with Aspose.Cells .NET | C# example for setting Shape.Line.DashStyle and Weight in Aspose.Cells | Apply custom line formatting to Excel shapes using Aspose.Cells API
// Tags: Aspose.Cells shape line dash formatting | Aspose.Cells shape border thickness setting | C# Excel shape line property modification | Aspose.Cells Shape.Line API usage | Excel shape custom line style Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load the workbook (use the provided load rule)
Workbook workbook = new Workbook("input.xlsx");

// Access the first worksheet (adjust index if needed)
Worksheet worksheet = workbook.Worksheets[0];

// Find the shape named "Divider"
Shape dividerShape = worksheet.Shapes["Divider"];
if (dividerShape != null)
{
    // Set line style to dashed
    dividerShape.Line.DashStyle = MsoLineDashStyle.Dash;
    
    // Set line thickness to two points
    dividerShape.Line.Weight = 2.0;
}

// Save the workbook (use the provided save rule)
workbook.Save("output.xlsx");
