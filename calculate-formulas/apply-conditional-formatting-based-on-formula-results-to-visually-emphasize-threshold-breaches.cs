using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingDemo
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
                cells[i, 0].PutValue(i * 20); // Values: 0,20,40,...,180
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a condition that highlights cells where the value exceeds 100
            // Using an expression type with a formula that evaluates to TRUE/FALSE
            int conditionIdx = fcs.AddCondition(
                FormatConditionType.Expression,
                OperatorType.None,
                "=A1>100",
                null);

            // Retrieve the created condition and set its visual style
            FormatCondition condition = fcs[conditionIdx];
            condition.Style.BackgroundColor = Color.Yellow; // Highlight background
            condition.Style.Font.Color = Color.Red;         // Highlight font color
            condition.StopIfTrue = true;                    // Stop further rules if this one applies

            // Save the workbook
            workbook.Save("ConditionalFormattingThreshold.xlsx");
        }
    }
}