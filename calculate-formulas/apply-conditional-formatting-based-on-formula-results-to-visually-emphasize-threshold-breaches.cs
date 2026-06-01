using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

class ConditionalFormattingDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1:A20)
            for (int i = 0; i < 20; i++)
            {
                worksheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,190
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A1:A20)
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = 19,
                StartColumn = 0,
                EndColumn = 0
            };
            conditions.AddArea(range);

            // Condition 1: Highlight cells with value greater than 100 (red background)
            int conditionIdx1 = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "100",
                null);
            FormatCondition condition1 = conditions[conditionIdx1];
            condition1.Style.BackgroundColor = Color.Red;

            // Condition 2: Highlight cells with value less than 50 (light green background)
            int conditionIdx2 = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessThan,
                "50",
                null);
            FormatCondition condition2 = conditions[conditionIdx2];
            condition2.Style.BackgroundColor = Color.LightGreen;

            // Save the workbook with the applied conditional formatting
            string outputPath = "ConditionalFormattingThreshold.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}