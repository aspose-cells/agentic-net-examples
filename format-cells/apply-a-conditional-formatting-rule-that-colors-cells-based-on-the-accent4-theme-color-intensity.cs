using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingAccent4Demo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                worksheet.Cells[i, 0].PutValue(i + 1);
            }

            // Add an empty conditional formatting collection
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a ColorScale condition
            fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition fc = fcs[0]; // The newly added condition

            // Configure the ColorScale: use Accent4 theme color with different tints for min and max
            // Minimum value: Accent4 darkened (tint = -0.5)
            CellsColor minColor = workbook.CreateCellsColor();
            minColor.ThemeColor = new ThemeColor(ThemeColorType.Accent4, -0.5);
            fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
            fc.ColorScale.MinColor = minColor.Color; // Resolve to actual Color

            // Maximum value: Accent4 lightened (tint = 0.5)
            CellsColor maxColor = workbook.CreateCellsColor();
            maxColor.ThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.5);
            fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
            fc.ColorScale.MaxColor = maxColor.Color; // Resolve to actual Color

            // Optional: set a middle point using the base Accent4 color (no tint)
            CellsColor midColor = workbook.CreateCellsColor();
            midColor.ThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.0);
            fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
            fc.ColorScale.MidCfvo.Value = 50; // 50th percentile (median)
            fc.ColorScale.MidColor = midColor.Color;

            // Save the workbook
            workbook.Save("ConditionalFormattingAccent4.xlsx");
        }
    }
}