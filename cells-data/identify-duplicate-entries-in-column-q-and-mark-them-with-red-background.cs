using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // OPTIONAL: populate column Q (index 16) with sample data
        // cells["Q1"].PutValue("Apple");
        // cells["Q2"].PutValue("Banana");
        // cells["Q3"].PutValue("Apple"); // duplicate
        // cells["Q4"].PutValue("Cherry");
        // cells["Q5"].PutValue("Banana"); // duplicate

        // Determine the last row that contains data in the sheet
        int lastRow = cells.MaxDataRow; // overall max data row

        // Define the range covering column Q (0‑based column index 16)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = lastRow,
            StartColumn = 16,
            EndColumn = 16
        };

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

        // Apply the range to the conditional formatting
        fcs.AddArea(range);

        // Add a condition that highlights duplicate values
        int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition condition = fcs[conditionIndex];

        // Create a style with a red background
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;

        // Assign the style to the condition
        condition.Style = redStyle;

        // Save the workbook (lifecycle rule)
        workbook.Save("DuplicateValuesInColumnQ.xlsx");
    }
}