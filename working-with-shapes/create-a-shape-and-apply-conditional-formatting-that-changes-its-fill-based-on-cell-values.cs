// Title: C# – Aspose.Cells: Create a rectangle shape and change its fill color based on cell values
// Description: C# example that adds numbers to B2:B4, defines conditional‑formatting rules, inserts a rectangle shape, and dynamically sets the shape’s fill to light green or light coral according to B4’s value, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | rectangle shape | shape fill color | conditional formatting | cell value driven color | Excel automation | sample code | GitHub example | dynamic fill | workbook
// Common Searches: Aspose.Cells change shape fill color based on cell value | C# conditional formatting shape color Aspose.Cells | How to bind rectangle fill to Excel cell in .NET | Dynamic shape color using Aspose.Cells | Example of shape fill with conditional formatting Aspose
// Developer Intent: Generate a workbook where a rectangle shape’s fill color updates automatically according to a cell’s numeric value using conditional formatting.
// Use Cases: KPI dashboard where a colored bar reflects performance metrics. | Risk matrix that highlights high‑risk items with a colored shape. | Automated report templates that adjust visual indicators when data changes. | Interactive Excel UI that uses shapes as status lights driven by cell values.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape, apply conditional formatting to B2:B4, and set the shape’s fill color based on the value in B4. | Explain how to use CellsColor.IsShapeColor to assign a dynamic fill to a shape in Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that demonstrates both cell background conditional formatting and synchronized shape fill colors.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that adds numbers to B2:B4, defines conditional‑formatting rules, inserts a rectangle shape, and dynamically sets the shape’s fill to light green or light coral according to B4’s value, using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample values that will drive the conditional formatting
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(70);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 0, 2, 100, 200, 100);
        shape.IsFilled = true;                         // Make the fill visible
        shape.Fill.FillType = FillType.Solid;          // Use solid fill
        shape.Fill.SolidFill.Color = Color.Gray;       // Default fill color

        // ---------- Conditional formatting on the cells ----------
        // Create a conditional formatting collection for the range B2:B4
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range B2:B4
        CellArea area = new CellArea
        {
            StartRow = 1,   // Row index is zero‑based (B2 -> row 1)
            EndRow = 3,
            StartColumn = 1, // Column B -> index 1
            EndColumn = 1
        };
        fcc.AddArea(area);

        // Condition 1: value > 50 → light green background
        int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
        FormatCondition fc = fcc[condIdx];
        fc.Style.BackgroundColor = Color.LightGreen;

        // Condition 2: value <= 50 → light coral background
        condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessOrEqual, "50", null);
        fc = fcc[condIdx];
        fc.Style.BackgroundColor = Color.LightCoral;

        // ---------- Apply shape fill based on a cell value ----------
        // For demonstration, use the value in B4 to decide the shape's fill color
        double cellValue = sheet.Cells["B4"].DoubleValue;
        CellsColor shapeColor = workbook.CreateCellsColor();
        shapeColor.IsShapeColor = true; // Indicate that this color is for a shape

        if (cellValue > 50)
        {
            shapeColor.Color = Color.LightGreen;
        }
        else
        {
            shapeColor.Color = Color.LightCoral;
        }

        // Assign the determined color to the shape's fill
        shape.Fill.SolidFill.Color = shapeColor.Color;

        // Save the workbook
        workbook.Save("ShapeConditionalFormattingDemo.xlsx");
    }
}
