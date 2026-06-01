using System.Drawing;
using Aspose.Cells;

class DynamicThresholdConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...
        }

        // Define a dynamic threshold value in cell B1
        worksheet.Cells["B1"].PutValue(35); // Threshold can be changed later

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea dataRange = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Associate the range with the conditional formatting
        fcc.AddArea(dataRange);

        // Add an expression condition: highlight cells where the cell value exceeds the threshold in B1
        // Formula uses absolute reference to B1 so the threshold is dynamic
        int conditionIndex = fcc.AddCondition(
            FormatConditionType.Expression,
            OperatorType.None,
            "=A1>$B$1",   // Formula1
            null);        // Formula2 not needed for Expression type

        // Retrieve the created condition and set its visual style
        FormatCondition condition = fcc[conditionIndex];
        condition.Style.BackgroundColor = Color.Yellow;

        // Save the workbook
        workbook.Save("DynamicThresholdDemo.xlsx");
    }
}