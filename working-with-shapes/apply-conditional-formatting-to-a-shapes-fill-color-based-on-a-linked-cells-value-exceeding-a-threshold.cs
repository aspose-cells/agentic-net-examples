// Title: C# Aspose.Cells Example: Change a Shape’s Fill Color via Conditional Formatting Linked to a Cell
// Description: Demonstrates how to create a workbook, insert a rectangle shape, link it to cell A1, set the shape’s fill to red, and add a conditional‑formatting rule that turns the cell background yellow when the value exceeds 50. The example shows how the linked cell drives visual formatting of the shape in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | shape conditional formatting | linked cell | rectangle shape fill color | CellsColor.IsShapeColor | Excel automation | conditional formatting rule | KPI dashboard visual cue | Aspose.Cells example
// Common Searches: Aspose.Cells change shape fill color based on cell value | C# link rectangle to worksheet cell Aspose.Cells | conditional formatting trigger shape color Aspose.Cells | how to use CellsColor.IsShapeColor in Aspose.Cells | Aspose.Cells example for KPI dashboard shapes
// Developer Intent: Automatically adjust a shape’s fill color when the value in its linked worksheet cell crosses a defined threshold.
// Use Cases: KPI dashboards where a status rectangle turns red if a metric exceeds its target. | Financial reports that highlight a shape when profit or loss surpasses a specified amount. | Interactive spreadsheets that synchronize visual alerts with cell‑driven thresholds.
// AI Prompts: Generate C# Aspose.Cells code that updates a rectangle’s fill color when the linked cell value is greater than 50. | Show how to set CellsColor.IsShapeColor and tie it to a conditional‑formatting rule on a worksheet cell. | Provide a complete example that links a shape to a cell and changes both the cell background and shape fill based on a numeric threshold.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConditionalFormattingDemo
{
    // Demonstrates how to create a workbook, insert a rectangle shape, link it to cell A1, set the shape’s fill to red, and add a conditional‑formatting rule that turns the cell background yellow when the value exceeds 50. The example shows how the linked cell drives visual formatting of the shape in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample numeric value in cell A1 (this cell will be linked to the shape)
            sheet.Cells["A1"].PutValue(30); // Change this value to test the condition

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, lower right row, lower right column, lower right offset
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 0, 5, 0, 0);
            rect.Name = "LinkedRectangle";

            // Link the shape to cell A1 so that the shape can read the cell's value
            rect.SetLinkedCell("A1", false, false);

            // Create a CellsColor instance and mark it as a shape color
            CellsColor shapeColor = workbook.CreateCellsColor();
            shapeColor.IsShapeColor = true;          // Indicates the color is intended for a shape
            shapeColor.Color = Color.Red;            // Desired fill color when the condition is met

            // Apply the color to the shape's fill (solid fill)
            rect.Fill.FillType = FillType.Solid;
            rect.Fill.SolidFill.Color = shapeColor.Color; // Use the Color property; IsShapeColor informs Aspose about the intent

            // -----------------------------------------------------------------
            // Add conditional formatting to the linked cell (A1)
            // The rule: if the cell value > 50, change the cell's background to Yellow
            // (The shape's fill color is already set; the conditional format demonstrates the link)
            // -----------------------------------------------------------------
            int cfIndex = sheet.ConditionalFormattings.Add();                     // Create a new conditional formatting collection
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Define the range that the conditional formatting will apply to (only A1)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 0,
                EndColumn = 0
            };
            conditions.AddArea(area);

            // Add a CellValue condition: value greater than 50
            int conditionIdx = conditions.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = conditions[conditionIdx];
            condition.Style.BackgroundColor = Color.Yellow; // Cell background when condition is true

            // -----------------------------------------------------------------
            // Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("ShapeConditionalFormattingDemo.xlsx");
        }
    }
}
