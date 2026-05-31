using System;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data that will be used by the conditional format
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(15);
            cells["A3"].PutValue(25);
            cells["A4"].PutValue(35);
            cells["A5"].PutValue(45);

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A5)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 4,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a condition: highlight cells with values between 10 and 30 (inclusive)
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "10", "30");
            FormatCondition condition = fcc[conditionIdx];

            // Set the style for the condition (red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = System.Drawing.Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // After setting up conditional formatting, recalculate formulas
            // (ensures any formulas that depend on the formatted cells are updated)
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ConditionalFormattingWithRecalc.xlsx");
        }
    }
}