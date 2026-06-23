using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalFormattingThemeAccent
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea range = CellArea.CreateCellArea(0, 0, 9, 0);

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Add the range to the collection
            fcc.AddArea(range);

            // Add a condition: cells with value greater than 50
            int conditionIdx = fcc.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                null);

            FormatCondition condition = fcc[conditionIdx];

            // Create a CellsColor instance and set its ThemeColor to Accent2 (no tint)
            CellsColor themeColor = workbook.CreateCellsColor();
            themeColor.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

            // Create a style and apply the theme color as the fill (foreground) color
            Style style = workbook.CreateStyle();
            style.ForegroundColor = themeColor.Color; // Resolve the actual Color from theme
            style.Pattern = BackgroundType.Solid;     // Ensure solid fill

            // Assign the style to the conditional format
            condition.Style = style;

            // Save the workbook
            workbook.Save("ConditionalFormatting_Accent2.xlsx");
        }
    }
}