using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in column A
            sheet.Cells["A1"].PutValue(30);
            sheet.Cells["A2"].PutValue(70);
            sheet.Cells["A3"].PutValue(45);

            // Add a rectangle shape
            // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 150, 200, 100);
            shape.IsFilled = true; // Ensure the shape has a fill

            // Apply conditional formatting to the range A1:A3
            // Cells with value > 50 will have a red background, otherwise green
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range for conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Condition: value > 50 -> red background
            int redCondIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition redCond = fcs[redCondIdx];
            redCond.Style.BackgroundColor = Color.Red;

            // Condition: value <= 50 -> green background
            int greenCondIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.LessOrEqual, "50", null);
            FormatCondition greenCond = fcs[greenCondIdx];
            greenCond.Style.BackgroundColor = Color.Green;

            // Determine shape fill based on the value of cell A2 (as an example)
            // If A2 > 50, set shape fill to red; otherwise, set it to green
            double cellValue = sheet.Cells["A2"].DoubleValue;
            CellsColor shapeColor = workbook.CreateCellsColor();
            shapeColor.IsShapeColor = true; // Indicate this color is for a shape

            if (cellValue > 50)
                shapeColor.Color = Color.Red;
            else
                shapeColor.Color = Color.Green;

            // Apply the determined color to the shape's fill
            shape.Fill.FillType = FillType.Solid;
            shape.Fill.SolidFill.Color = shapeColor.Color;

            // Save the workbook
            workbook.Save("ShapeConditionalFormattingDemo.xlsx");
        }
    }
}