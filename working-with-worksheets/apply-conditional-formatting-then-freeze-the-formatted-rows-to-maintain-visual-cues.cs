using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingAndFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (rows 0-9, columns A-C)
            for (int row = 0; row < 10; row++)
            {
                worksheet.Cells[row, 0].PutValue($"Item {row + 1}");
                worksheet.Cells[row, 1].PutValue(row * 10);          // numeric value for conditional formatting
                worksheet.Cells[row, 2].PutValue(DateTime.Today.AddDays(row));
            }

            // ---------- Apply Conditional Formatting ----------
            // Add a new conditional formatting rule collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the rule will apply (rows 0-4, columns B)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 1,
                EndColumn = 1
            };
            cfCollection.AddArea(area);

            // Add a condition: highlight cells with value > 50
            int conditionIdx = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                null);

            // Set the style for the condition (yellow background)
            FormatCondition condition = cfCollection[conditionIdx];
            condition.Style.BackgroundColor = Color.Yellow;

            // ---------- Freeze the formatted rows ----------
            // Freeze the first 5 rows so that the conditional formatting remains visible while scrolling
            // Freeze at cell A6 (row index 5) with 5 frozen rows and 0 frozen columns
            worksheet.FreezePanes("A6", 5, 0);

            // Save the workbook
            workbook.Save("ConditionalFormattingAndFreeze.xlsx");
        }
    }
}