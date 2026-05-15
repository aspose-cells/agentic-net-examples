using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    class Program
    {
        static void Main()
        {
            // Threshold value for highlighting
            double threshold = 50.0;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column A (rows 0-9)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a condition: cell value greater than the threshold
            int conditionIndex = fcc.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                null);

            // Retrieve the condition and set its style (e.g., light orange background)
            FormatCondition condition = fcc[conditionIndex];
            condition.Style.BackgroundColor = Color.Orange;
            condition.Style.Font.Color = Color.Black;

            // Save the workbook
            workbook.Save("ConditionalFormatting_GreaterThanThreshold.xlsx");
        }
    }
}