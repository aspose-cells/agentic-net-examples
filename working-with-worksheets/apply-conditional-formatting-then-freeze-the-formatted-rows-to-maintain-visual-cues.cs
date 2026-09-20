// Title: Add red background conditional formatting for values > 50 in column A and freeze the first ten rows using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that sets a red fill for cells in column A whose numeric value exceeds 50, then locks rows 1‑10 with FreezePanes. | Create an Aspose.Cells example that defines a cell‑value based style rule for A1:A10 and freezes the worksheet pane so the formatted rows stay visible.
// Common Searches: Aspose.Cells how to apply a red background conditional format to a column based on value threshold | C# freeze top ten rows after adding conditional formatting with Aspose.Cells | example of using ConditionalFormattings and FreezePanes together in Aspose.Cells .NET
// Tags: format rule red background Aspose.Cells C# | pane freezing top ten rows Aspose.Cells | cell value greater than 50 style condition Aspose.Cells | format collection usage with pane freezing Aspose.Cells | Aspose.Cells workbook visual cue preservation

using Aspose.Cells;
using System;
using System.Drawing;

// The program creates a workbook, fills column A with numbers, adds a conditional formatting rule that colors cells with values greater than 50 red, freezes the first ten rows to keep the visual cues visible, and saves the file as ConditionalFormattingAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1-10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // A1:A10
            }

            // Define the range for conditional formatting (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a conditional formatting rule: highlight cells > 50 with red background
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(area);
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = cfCollection[conditionIndex];
            condition.Style = workbook.CreateStyle();
            condition.Style.ForegroundColor = Color.Red;
            condition.Style.Pattern = BackgroundType.Solid;

            // Freeze rows 1‑10 (freeze panes above row 11)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(11, 0, 10, 0);

            // Save the workbook
            workbook.Save("ConditionalFormattingAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
