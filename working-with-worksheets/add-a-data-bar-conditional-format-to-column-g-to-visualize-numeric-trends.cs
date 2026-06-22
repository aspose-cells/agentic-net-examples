using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column G (index 6) with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 6].PutValue(i * 10); // G1:G10 = 0,10,20,...,90
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range for the data bar (column G, rows 0-9)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 6,
            EndColumn = 6
        };
        fcs.AddArea(area);

        // Add a DataBar condition to the range
        int conditionIndex = fcs.AddCondition(FormatConditionType.DataBar);
        FormatCondition condition = fcs[conditionIndex];

        // Configure the DataBar properties
        DataBar dataBar = condition.DataBar;
        dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
        dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
        dataBar.Color = Color.LightBlue;   // Visual color for the data bar
        dataBar.ShowValue = true;          // Show cell values alongside the bar

        // Save the workbook
        workbook.Save("DataBarColumnG.xlsx", SaveFormat.Xlsx);
    }
}