using System;
using System.Drawing;
using Aspose.Cells;

class ValidateConditionalFormattingTheme
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (0,10,20,30,40)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 10);
        }

        // Add conditional formatting that uses a theme color (Accent2) for the font
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A5
        CellArea area = new CellArea { StartRow = 0, EndRow = 4, StartColumn = 0, EndColumn = 0 };
        fcc.AddArea(area);

        // Condition: cell value > 15
        int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "15", null);
        FormatCondition fc = fcc[condIdx];

        // Set the style to use theme color Accent2
        Style cfStyle = workbook.CreateStyle();
        cfStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
        fc.Style = cfStyle;

        // Change the theme: set Accent2 to a distinct color (Orange)
        Color newAccent2 = Color.Orange;
        workbook.SetThemeColor(ThemeColorType.Accent2, newAccent2);

        // Retrieve a cell that satisfies the condition (A3 = 20)
        Cell testCell = sheet.Cells[2, 0];

        // Get the conditional formatting result after the theme change
        ConditionalFormattingResult result = testCell.GetConditionalFormattingResult();

        // Extract the resolved font color from the conditional style
        Color resolvedColor = result?.ConditionalStyle?.Font?.Color ?? Color.Empty;

        // Validate that the resolved color matches the new theme color
        bool isCorrect = resolvedColor.ToArgb() == newAccent2.ToArgb();

        Console.WriteLine($"Resolved font color: {resolvedColor}");
        Console.WriteLine($"Expected theme color: {newAccent2}");
        Console.WriteLine($"Validation {(isCorrect ? "passed" : "failed")}.");

        // Save the workbook (optional)
        workbook.Save("ThemeValidation.xlsx");
    }
}