using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample numeric data in column A (rows 2‑11)
        for (int i = 0; i < 10; i++)
        {
            // Values: 5, 15, 25, …, 95
            cells[i + 1, 0].PutValue(i * 10 + 5);
        }

        // Merge cells A1:B1 to create a header area
        cells.Merge(0, 0, 1, 2);
        cells[0, 0].PutValue("Performance");

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range that will be conditionally formatted (A2:A11)
        CellArea dataArea = new CellArea
        {
            StartRow = 1,
            EndRow = 10,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(dataArea);

        // 1) Values between 40 and 80 → Yellow background
        int condBetween = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "40", "80");
        FormatCondition fcBetween = fcc[condBetween];
        fcBetween.Style.BackgroundColor = Color.Yellow;

        // 2) Values greater than 80 → Green background, white bold font
        int condGreater = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "80", null);
        FormatCondition fcGreater = fcc[condGreater];
        fcGreater.Style.BackgroundColor = Color.Green;
        fcGreater.Style.Font.Color = Color.White;
        fcGreater.Style.Font.IsBold = true;

        // 3) Values less than 40 → Red background
        int condLess = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "40", null);
        FormatCondition fcLess = fcc[condLess];
        fcLess.Style.BackgroundColor = Color.Red;

        // Save the workbook with merged header and conditional formatting applied
        workbook.Save("MergedConditionalFormatting.xlsx");
    }
}