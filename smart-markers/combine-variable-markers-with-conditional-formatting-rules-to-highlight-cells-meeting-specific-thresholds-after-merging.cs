// Title: How to merge a header row and apply three value‑based conditional formatting rules to a column using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that merges cells A1:B1, writes a title, populates A2:A10 with numbers, and creates three conditional formatting rules: values >70 red, 40‑70 yellow, <40 green. | Demonstrate saving the workbook as an XLS file with the MergeAreas option enabled after the conditional formats are applied.
// Common Searches: Aspose.Cells C# merge cells and add conditional formatting based on numeric thresholds | example of multiple cell value conditions (greater than, between, less than) in Aspose.Cells | how to enable MergeAreas when saving an XLS workbook with Aspose.Cells | apply red, yellow, green background colors to an Excel column using Aspose.Cells conditional formatting | C# code to create a header row and conditional formatting after merging cells in Aspose.Cells
// Tags: Aspose.Cells merge header cells C# | Aspose.Cells conditional formatting multiple rules | conditional formatting cellvalue thresholds Aspose.Cells | Save workbook with MergeAreas option Aspose.Cells | C# Excel XLS output Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingAfterMerge
{
    // The example creates a workbook, merges A1:B1 as a centered header, fills A2:A10 with values 10‑90, and adds three conditional formatting rules (red for >70, yellow for 40‑70, light green for <40) to that column. The file is saved as an XLS with MergeAreas enabled to handle merged cells efficiently.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (rows 2 to 10)
            for (int i = 1; i <= 9; i++) // zero‑based index, row 1 = A2
            {
                cells[i, 0].PutValue(i * 10); // Values: 10,20,...,90
            }

            // Merge header cells A1:B1 and set a title
            cells.Merge(0, 0, 1, 2);
            cells[0, 0].PutValue("Sales Data");
            Style headerStyle = workbook.CreateStyle();
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 14;
            cells[0, 0].SetStyle(headerStyle);

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A2:A10)
            CellArea dataArea = new CellArea
            {
                StartRow = 1,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(dataArea);

            // Condition 1: Values greater than 70 -> Red background
            int condIdx1 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "70", null);
            FormatCondition fc1 = fcc[condIdx1];
            fc1.Style.BackgroundColor = Color.Red;
            fc1.Style.Font.Color = Color.White;

            // Condition 2: Values between 40 and 70 -> Yellow background
            int condIdx2 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "40", "70");
            FormatCondition fc2 = fcc[condIdx2];
            fc2.Style.BackgroundColor = Color.Yellow;
            fc2.Style.Font.Color = Color.Black;

            // Condition 3: Values less than 40 -> Green background
            int condIdx3 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "40", null);
            FormatCondition fc3 = fcc[condIdx3];
            fc3.Style.BackgroundColor = Color.LightGreen;
            fc3.Style.Font.Color = Color.Black;

            // Save the workbook with MergeAreas enabled to optimize merged cells handling
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            saveOptions.MergeAreas = true;
            workbook.Save("ConditionalFormattingAfterMerge.xls", saveOptions);
        }
    }
}
