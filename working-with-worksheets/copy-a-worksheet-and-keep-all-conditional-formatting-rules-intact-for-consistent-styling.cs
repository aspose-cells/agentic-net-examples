using System.Drawing;
using Aspose.Cells;

class WorksheetCopyWithConditionalFormatting
{
    static void Main()
    {
        // Create source workbook and get its first worksheet
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate some sample data
        sourceSheet.Cells["A1"].PutValue(5);
        sourceSheet.Cells["A2"].PutValue(15);
        sourceSheet.Cells["A3"].PutValue(25);

        // Add conditional formatting: cells > 10 get a red background
        int cfIndex = sourceSheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sourceSheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A3 for the formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add the condition and style
        int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        FormatCondition condition = fcc[conditionIdx];
        Style redStyle = sourceWorkbook.CreateStyle();
        redStyle.BackgroundColor = Color.Red;
        condition.Style = redStyle;

        // Create destination workbook and get its first worksheet
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];
        destSheet.Name = "Destination";

        // Copy contents and formats from source to destination worksheet
        destSheet.Copy(sourceSheet);

        // Explicitly copy conditional formatting collection to ensure it is retained
        destSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

        // Save the resulting workbook
        destWorkbook.Save("WorksheetCopyWithConditionalFormatting.xlsx");
    }
}