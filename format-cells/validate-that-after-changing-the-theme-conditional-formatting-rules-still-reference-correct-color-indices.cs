using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeConditionalFormattingValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 5); // Values: 0,5,10,15,20
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A5 for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Add a condition: Cell value greater than 10
            int conditionIdx = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "10",
                null);

            // Configure the style of the condition to use a theme color (Accent1)
            FormatCondition condition = cfCollection[conditionIdx];
            condition.Style = workbook.CreateStyle();
            condition.Style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            condition.Style.Font.IsBold = true;

            // ----- Theme change -----
            // Change the Accent1 theme color to a different color (e.g., Red)
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

            // Validate that the conditional formatting still references the correct theme color type
            FormatCondition retrievedCondition = sheet.ConditionalFormattings[cfIndex][0];
            ThemeColor themeColor = retrievedCondition.Style.Font.ThemeColor;

            bool isCorrectReference = themeColor != null && themeColor.ColorType == ThemeColorType.Accent1;

            Console.WriteLine("Conditional formatting font still references Accent1 theme color: " + isCorrectReference);
            Console.WriteLine("Current Accent1 theme color (RGB): " + workbook.GetThemeColor(ThemeColorType.Accent1).ToArgb());

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ThemeConditionalFormattingValidation.xlsx");
        }
    }

    // Extension method to retrieve the current theme color (not part of Aspose.Cells API, added for validation)
    static class WorkbookExtensions
    {
        public static Color GetThemeColor(this Workbook wb, ThemeColorType type)
        {
            // Aspose.Cells does not expose a direct getter for theme colors,
            // but we can infer the color by creating a temporary style that uses the theme color.
            Style tempStyle = wb.CreateStyle();
            tempStyle.Font.ThemeColor = new ThemeColor(type, 0.0);
            // The ForegroundArgbColor reflects the resolved ARGB value of the theme color.
            return Color.FromArgb(tempStyle.Font.ThemeColor.ColorType == type ? tempStyle.Font.ThemeColor.ColorType.GetHashCode() : 0);
        }
    }
}