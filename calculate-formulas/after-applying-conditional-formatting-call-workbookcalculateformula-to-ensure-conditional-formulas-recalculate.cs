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

            // Populate sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["B1"].Formula = "=A1*2";   // Formula that will be affected by conditional formatting
            cells["B2"].Formula = "=A2*2";
            cells["B3"].Formula = "=A3*2";

            // Add a conditional formatting rule: highlight cells in column A between 15 and 25
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A3 for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 2,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add the condition (type: CellValue, operator: Between, formulas "15" and "25")
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "15", "25");
            FormatCondition condition = fcc[conditionIdx];

            // Set the style to be applied when the condition is met (e.g., red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = System.Drawing.Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // After setting up conditional formatting, recalculate all formulas
            // This ensures that any formulas dependent on the formatted cells are updated
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ConditionalFormattingWithRecalc.xlsx");
        }
    }
}