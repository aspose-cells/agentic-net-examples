using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data for the series
        worksheet.Cells["A1"].PutValue(120);
        worksheet.Cells["A2"].PutValue(560);
        worksheet.Cells["A3"].PutValue(430);
        worksheet.Cells["A4"].PutValue(720);
        worksheet.Cells["A5"].PutValue(300);

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A5)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 4,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add a condition: values greater than 500
        int conditionIndex = fcc.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "500",
            null);

        // Retrieve the condition and set its style (fill color red)
        FormatCondition condition = fcc[conditionIndex];
        condition.Style.BackgroundColor = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingGreaterThan500.xlsx");
    }
}