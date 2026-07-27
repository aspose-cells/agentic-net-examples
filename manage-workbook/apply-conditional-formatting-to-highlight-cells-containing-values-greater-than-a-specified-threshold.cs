using System;
using System.Drawing;
using Aspose.Cells;

class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (rows 1-10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
        }

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the range with the conditional formatting
        fcc.AddArea(range);

        // Specify the threshold value
        const string threshold = "50";

        // Add a condition: highlight cells with values greater than the threshold
        int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, threshold, null);
        FormatCondition condition = fcc[conditionIdx];

        // Set the style to be applied when the condition is met (e.g., light orange background)
        condition.Style.BackgroundColor = Color.Orange;
        condition.Style.Font.Color = Color.Black;

        // Save the workbook
        workbook.Save("ConditionalFormatting_GreaterThanThreshold.xlsx");
    }
}