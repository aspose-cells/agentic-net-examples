using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsErrorHighlightDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with formulas that will produce errors
            cells["A1"].PutValue(10);
            cells["A2"].Formula = "=A1/0";               // #DIV/0! error
            cells["A3"].Formula = "=UNKNOWNFUNC(1)";    // #NAME? error
            cells["A4"].Formula = "=INDIRECT(\"Z1000\")"; // #REF! error

            // Calculate all formulas so error values are materialized
            workbook.CalculateFormula();

            // Add a conditional formatting rule that highlights cells containing errors
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the rule applies (entire used range)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = cells.MaxDataRow,
                EndColumn = cells.MaxDataColumn
            };
            cfCollection.AddArea(area);

            // Add the "ContainsErrors" condition
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.ContainsErrors);
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the highlight style (e.g., yellow background)
            condition.Style.BackgroundColor = Color.Yellow;

            // Save the workbook (lifecycle: save)
            workbook.Save("ErrorHighlighted.xlsx");
        }
    }
}