using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the workbook that contains the data.
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet (adjust if needed).
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last row that contains data in the sheet.
        int lastRow = sheet.Cells.MaxDataRow; // zero‑based index

        // Define the range that covers column Q (index 16) from the first row to the last data row.
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = lastRow,
            StartColumn = 16,   // Column Q
            EndColumn = 16
        };

        // Add a new conditional formatting collection to the worksheet.
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Apply the defined area to the conditional formatting.
        fcs.AddArea(area);

        // Add a condition that highlights duplicate values.
        int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition condition = fcs[conditionIndex];

        // Create a style with a solid red background.
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;

        // Assign the style to the condition.
        condition.Style = redStyle;

        // Save the workbook with the applied formatting.
        workbook.Save("output.xlsx");
    }
}