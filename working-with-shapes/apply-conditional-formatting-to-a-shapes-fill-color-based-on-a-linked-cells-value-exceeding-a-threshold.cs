// Title: Aspose.Cells .NET – Change a Shape’s Fill Color Using Conditional Formatting Linked to a Cell
// Description: Demonstrates how to create a workbook, write a numeric value to a cell, link a rectangle shape to that cell, apply a conditional‑formatting rule (value > 50 → red), evaluate the rule, and set the shape’s fill color to match the cell’s background (red or default gray) before saving the file.
// Keywords: Aspose.Cells conditional formatting shape fill | link shape to cell Aspose.Cells | shape fill color based on cell value .NET | Aspose.Cells C# shape color programmatically | Excel shape conditional formatting Aspose | CellsColor.IsShapeColor | dynamic dashboard shapes Aspose.Cells
// Common Searches: Aspose.Cells change shape fill color based on cell value | conditional formatting affect shape color Aspose.Cells .NET | link Excel shape to cell and update fill color | apply conditional formatting to shape fill using Aspose | C# Aspose.Cells shape color from cell condition
// Developer Intent: Automatically update a shape’s fill color when the value in its linked cell meets a specified threshold.
// Use Cases: Progress‑bar rectangle that turns red when sales exceed a target. | Warning icon that changes to red when a KPI falls below a critical level. | Dynamic dashboard where multiple shapes reflect calculation results via color changes.
// AI Prompts: Generate C# code with Aspose.Cells to link a shape to a cell and sync its fill color to a conditional‑formatting rule. | Show how to use CellsColor.IsShapeColor to apply a cell’s conditional‑formatting result to a shape’s fill. | Explain step‑by‑step how to evaluate a cell’s conditional formatting outcome and mirror the color on a linked shape in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConditionalFormattingDemo
{
    // Demonstrates how to create a workbook, write a numeric value to a cell, link a rectangle shape to that cell, apply a conditional‑formatting rule (value > 50 → red), evaluate the rule, and set the shape’s fill color to match the cell’s background (red or default gray) before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample numeric value into cell B2 (row 1, column 1)
            Cell linkedCell = sheet.Cells["B2"];
            linkedCell.PutValue(75); // Example value exceeding the threshold

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            RectangleShape shape = sheet.Shapes.AddRectangle(2, 0, 2, 100, 100, 100);
            shape.Name = "ThresholdShape";

            // Link the shape to cell B2
            shape.SetLinkedCell("B2", false, false);

            // Define a conditional formatting rule for the linked cell:
            // If the cell value > 50, set the cell background to Red
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Apply the rule to cell B2 only
            CellArea area = new CellArea
            {
                StartRow = linkedCell.Row,
                EndRow = linkedCell.Row,
                StartColumn = linkedCell.Column,
                EndColumn = linkedCell.Column
            };
            fcs.AddArea(area);

            // Add the condition (CellValue > 50)
            int conditionIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = fcs[conditionIdx];
            condition.Style.BackgroundColor = Color.Red;

            // Recalculate formulas (if any) to ensure conditional formatting is evaluated
            workbook.CalculateFormula();

            // After calculation, inspect the cell's displayed background color.
            // If the background color is Red (i.e., condition met), apply the same color to the shape's fill.
            // Use CellsColor with IsShapeColor = true to indicate the color is intended for a shape.
            CellsColor shapeColor = workbook.CreateCellsColor();
            shapeColor.IsShapeColor = true;

            // Determine if the condition was met by checking the cell's style background color
            if (linkedCell.GetStyle().ForegroundColor.Equals(Color.Red.ToArgb()))
            {
                shapeColor.Color = Color.Red;
            }
            else
            {
                shapeColor.Color = Color.LightGray; // Default fill when condition not met
            }

            // Apply the determined color to the shape's fill
            shape.Fill.FillType = FillType.Solid;
            shape.Fill.SolidFill.Color = shapeColor.Color; // Use the Color property; IsShapeColor informs Aspose internally

            // Save the workbook
            workbook.Save("ShapeConditionalFormattingDemo.xlsx");
        }
    }
}
