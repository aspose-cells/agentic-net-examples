using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea range = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            EndRow = 9,        // Row 10
            StartColumn = 0,   // Column A
            EndColumn = 0      // Column A
        };
        conditions.AddArea(range);

        // Add a "contains text" condition
        int conditionIndex = conditions.AddCondition(FormatConditionType.ContainsText);
        FormatCondition condition = conditions[conditionIndex];

        // Set the text to look for and the formatting (red font color)
        condition.Text = "Error";
        condition.Style.Font.Color = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingError.xlsx");
    }
}