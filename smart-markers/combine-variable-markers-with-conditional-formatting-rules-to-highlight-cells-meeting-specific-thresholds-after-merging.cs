using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample numeric data in column A (rows 0‑9)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
        }

        // Merge the header cells A1:B1 (row 0, column 0, 1 row, 2 columns)
        cells.Merge(0, 0, 1, 2);
        cells[0, 0].PutValue("Performance");

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the formatting will be applied (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Condition 1: values between 30 and 70 → yellow background
        int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "30", "70");
        FormatCondition condition = fcc[conditionIdx];
        condition.Style.BackgroundColor = Color.Yellow;

        // Condition 2: values greater than 70 → green background, white bold font
        conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "70", null);
        condition = fcc[conditionIdx];
        condition.Style.BackgroundColor = Color.Green;
        condition.Style.Font.Color = Color.White;
        condition.Style.Font.IsBold = true;

        // Save the workbook with merged areas optimization enabled
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.MergeAreas = true;
        workbook.Save("MergedConditionalFormatting.xls", saveOptions);
    }
}