using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value in cell B2 (row 1, column 1)
                sheet.Cells["B2"].PutValue(75);

                // Add a rectangle shape and link it to cell B2
                Shape shape = sheet.Shapes.AddRectangle(2, 1, 2, 150, 100, 100);
                shape.SetLinkedCell("B2", true, true);

                // Create a shape color (must be marked as a shape color)
                CellsColor shapeColor = workbook.CreateCellsColor();
                shapeColor.IsShapeColor = true;
                shapeColor.Color = Color.LightGray; // default fill

                shape.Fill.FillType = FillType.Solid;
                shape.Fill.SolidFill.Color = shapeColor.Color;

                // Add conditional formatting to the linked cell (B2)
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

                // Define the range that the conditional formatting applies to (only B2)
                CellArea area = new CellArea
                {
                    StartRow = 1,   // zero‑based index for row 2
                    EndRow = 1,
                    StartColumn = 1, // zero‑based index for column B
                    EndColumn = 1
                };
                fcs.AddArea(area);

                // Add a condition: cell value greater than 80
                int conditionIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", null);
                FormatCondition condition = fcs[conditionIdx];

                // When the condition is true, change the shape fill color to Red
                // The style's background color is applied to the linked shape because IsShapeColor is true
                condition.Style.BackgroundColor = Color.Red;
                condition.Style.Pattern = BackgroundType.Solid; // ensure background color is used

                // Save the workbook
                string outputPath = "ShapeConditionalFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}