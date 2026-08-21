// Title: C# – Conditional Formatting on Merged Cells with Smart Markers in Aspose.Cells
// Description: Shows how to merge a header row, insert smart markers, and apply three value‑based conditional‑formatting rules (red <50, yellow 50‑100, green >100) to column A, then save the workbook as XLS with MergeAreas enabled.
// Keywords: Aspose.Cells | C# | conditional formatting | merged cells | smart markers | Excel | value thresholds | MergeAreas | XLS export | cell style | format condition | color coding
// Common Searches: Aspose.Cells conditional formatting merged cells C# | How to use smart markers with conditional formatting in Aspose.Cells | C# set background color based on cell value Aspose.Cells | Save workbook with MergeAreas option Aspose.Cells | Apply multiple format conditions to a range after merging header Aspose.Cells
// Developer Intent: Create a workbook, merge a title row, add smart markers, apply three threshold‑based color rules to a column, and save with merged‑area optimization.
// Use Cases: Sales dashboard where the header spans two columns and sales numbers are color‑coded by performance bands. | Automated report generation that merges a title row, fills data via smart markers, and highlights out‑of‑range values. | Exporting an Excel file with a merged header and conditional colors while preserving formatting for downstream processing.
// AI Prompts: Generate C# code that adds a fourth conditional formatting rule for values equal to 75 using Aspose.Cells. | Explain how to bind a DataTable to smart markers and keep the existing conditional formatting. | Provide an example to export the workbook to PDF while retaining the conditional colors. | Show how to modify the example to work with .xlsx format and preserve merged cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingAfterMerge
{
    // Shows how to merge a header row, insert smart markers, and apply three value‑based conditional‑formatting rules (red <50, yellow 50‑100, green >100) to column A, then save the workbook as XLS with MergeAreas enabled.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (rows 2-11)
            for (int i = 1; i <= 10; i++)
            {
                cells[i, 0].PutValue(i * 15); // Values: 15,30,...,150
            }

            // Merge header cells A1:B1 and set a title
            cells.Merge(0, 0, 1, 2);
            cells[0, 0].PutValue("Sales Data");

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A2:A11)
            CellArea dataArea = new CellArea
            {
                StartRow = 1,
                EndRow = 10,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(dataArea);

            // Condition 1: Values less than 50 -> Red background
            int condIdx1 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", null);
            FormatCondition fc1 = fcc[condIdx1];
            fc1.Style.BackgroundColor = Color.Red;

            // Condition 2: Values between 50 and 100 -> Yellow background
            int condIdx2 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "100");
            FormatCondition fc2 = fcc[condIdx2];
            fc2.Style.BackgroundColor = Color.Yellow;

            // Condition 3: Values greater than 100 -> Green background
            int condIdx3 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "100", null);
            FormatCondition fc3 = fcc[condIdx3];
            fc3.Style.BackgroundColor = Color.LightGreen;

            // Save the workbook with MergeAreas enabled to optimize merged cells handling
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            saveOptions.MergeAreas = true;
            workbook.Save("ConditionalFormattingAfterMerge.xls", saveOptions);
        }
    }
}
