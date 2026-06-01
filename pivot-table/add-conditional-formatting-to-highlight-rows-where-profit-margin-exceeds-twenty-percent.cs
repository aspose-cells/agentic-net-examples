using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample header
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["C1"].PutValue("Cost");
        sheet.Cells["D1"].PutValue("Margin");

        // Sample data rows (margin = (Revenue-Cost)/Revenue)
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["C2"].PutValue(800);
        sheet.Cells["D2"].PutValue(0.2); // 20%

        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["C3"].PutValue(1000);
        sheet.Cells["D3"].PutValue(0.3333); // 33.33%

        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(800);
        sheet.Cells["C4"].PutValue(700);
        sheet.Cells["D4"].PutValue(0.125); // 12.5%

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the formatting applies (rows 2‑100, columns A‑D)
        CellArea area = new CellArea
        {
            StartRow = 1,      // zero‑based index, row 2 in Excel
            EndRow = 99,       // row 100 in Excel
            StartColumn = 0,   // column A
            EndColumn = 3      // column D
        };
        fcc.AddArea(area);

        // Add an expression‑type condition: highlight rows where margin > 20%
        int condIndex = fcc.AddCondition(FormatConditionType.Expression);
        FormatCondition condition = fcc[condIndex];
        // Use a relative row reference; $D1 refers to column D of the current row
        condition.Formula1 = "=($D1>0.2)";

        // Define the style to apply (yellow fill)
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;
        highlightStyle.Pattern = BackgroundType.Solid;
        condition.Style = highlightStyle;

        // Save the workbook
        workbook.Save("ProfitMarginConditionalFormatting.xlsx");
    }
}