using System;
using Aspose.Cells;
using System.Drawing;

class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Get the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample numeric data in column A (A1:A10).
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
        }

        // Define the cell area to which the conditional formatting will be applied.
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Add a new conditional formatting collection to the worksheet.
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the defined range with the collection.
        fcc.AddArea(range);

        // Specify the threshold value.
        double threshold = 50;

        // Add a CellValue condition: highlight cells with values greater than the threshold.
        int conditionIdx = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            threshold.ToString(),
            null);
        FormatCondition condition = fcc[conditionIdx];

        // Set the formatting style (e.g., light green background).
        condition.Style.BackgroundColor = Color.LightGreen;

        // Save the workbook to a file.
        workbook.Save("ConditionalFormattingGreaterThan.xlsx");
    }
}