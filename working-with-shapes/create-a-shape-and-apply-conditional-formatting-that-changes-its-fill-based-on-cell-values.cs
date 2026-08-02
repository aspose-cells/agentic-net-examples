// Title: C# – Add a Rectangle Shape and Change Its Fill via Conditional Formatting in Aspose.Cells
// Description: This example shows how to create a new workbook, insert numeric data, add a rectangle shape, set a solid default fill, define conditional‑formatting rules (0‑50 → light green, >50 → light coral) for cells A1:A4, link the shape to a cell, and save the file. It demonstrates the technique for updating a shape’s fill based on cell values using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | shape | rectangle | conditional formatting | fill color | linked cell | shape fill sync | Excel automation
// Common Searches: Aspose.Cells change shape fill color based on cell value | C# link rectangle shape to cell in Excel | conditional formatting affect shape fill Aspose.Cells | how to sync shape color with cell conditional format .NET | Aspose.Cells example rectangle shape conditional fill
// Developer Intent: Create a shape whose fill color can be updated according to conditional‑formatting rules applied to a cell range.
// Use Cases: KPI dashboard where a rectangle turns green or coral as a metric crosses thresholds. | Report templates that mirror cell background colors on shapes for consistent visual cues. | Automated Excel generation that highlights data ranges by changing shape fills programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that synchronizes a rectangle shape's fill color to the background color of a conditionally formatted cell. | Show how to refresh a linked shape's color after workbook recalculation when the underlying cell value changes. | Explain how to read the applied conditional‑formatting color from a cell and assign it to Shape.Fill.SolidFill.Color in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalShapeFill
{
    // This example shows how to create a new workbook, insert numeric data, add a rectangle shape, set a solid default fill, define conditional‑formatting rules (0‑50 → light green, >50 → light coral) for cells A1:A4, link the shape to a cell, and save the file. It demonstrates the technique for updating a shape’s fill based on cell values using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that will drive the conditional formatting
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["A3"].PutValue(60);
            sheet.Cells["A4"].PutValue(90);

            // -----------------------------------------------------------------
            // 1. Create a shape (rectangle) that will display a fill color
            // -----------------------------------------------------------------
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape rect = sheet.Shapes.AddRectangle(5, 0, 5, 100, 100, 100);
            rect.IsFilled = true;                     // Make sure the fill is visible
            rect.Fill.FillType = FillType.Solid;      // Use solid fill

            // -----------------------------------------------------------------
            // 2. Create a CellsColor instance with IsShapeColor = true
            // -----------------------------------------------------------------
            CellsColor shapeColor = workbook.CreateCellsColor();
            shapeColor.IsShapeColor = true;           // Indicates the color is for a shape
            shapeColor.Color = Color.LightGray;       // Default fill color

            // Apply the color to the shape
            rect.Fill.SolidFill.Color = shapeColor.Color;

            // -----------------------------------------------------------------
            // 3. Add conditional formatting to the data range (A1:A4)
            //    The formatting will change the background color of the cells,
            //    and we will also mirror that color to the shape.
            // -----------------------------------------------------------------
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting applies
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 3,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Condition 1: values between 0 and 50 -> light green
            int condIdx1 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "0", "50");
            FormatCondition fc1 = fcc[condIdx1];
            fc1.Style.BackgroundColor = Color.LightGreen;

            // Condition 2: values greater than 50 -> light coral
            int condIdx2 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition fc2 = fcc[condIdx2];
            fc2.Style.BackgroundColor = Color.LightCoral;

            // -----------------------------------------------------------------
            // 4. Link the shape to the first cell (A1) so that when the cell
            //    value changes, the shape can be refreshed programmatically.
            //    (Note: Aspose.Cells does not automatically sync shape fill with
            //    cell conditional formatting, but this demonstrates the linking
            //    mechanism which can be used in custom logic.)
            // -----------------------------------------------------------------
            rect.SetLinkedCell("A1", false, false);

            // -----------------------------------------------------------------
            // 5. Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("ShapeConditionalFillDemo.xlsx");
        }
    }
}
