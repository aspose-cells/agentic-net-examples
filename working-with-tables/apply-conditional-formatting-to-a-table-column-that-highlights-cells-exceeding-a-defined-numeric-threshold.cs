using Aspose.Cells;
using System.Drawing;

class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample numeric data in column B (index 1)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 1].PutValue(i * 10); // 0,10,20,...,90
        }

        // Define the range that covers the entire column B (rows 0‑9)
        CellArea columnRange = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the range with the conditional formatting collection
        fcc.AddArea(columnRange);

        // Define the numeric threshold
        double threshold = 50;

        // Add a condition: highlight cells where the value is greater than the threshold
        int conditionIndex = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            threshold.ToString(),
            null);

        // Retrieve the condition and set its formatting (yellow background)
        FormatCondition condition = fcc[conditionIndex];
        condition.Style.BackgroundColor = Color.Yellow;

        // Save the workbook to a file
        workbook.Save("ConditionalFormattingResult.xlsx");
    }
}