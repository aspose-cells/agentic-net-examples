using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i);
        }

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcs.AddArea(area);

        // Add a condition: cells with values greater than 5
        int conditionIndex = fcs.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "5",
            null);
        FormatCondition fc = fcs[conditionIndex];

        // Apply a style that uses the Accent4 theme color with a tint (intensity) of 0.3
        fc.Style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.3);
        fc.Style.Pattern = BackgroundType.Solid; // Ensure the foreground color is visible

        // Save the workbook
        workbook.Save("ConditionalAccent4.xlsx");
    }
}