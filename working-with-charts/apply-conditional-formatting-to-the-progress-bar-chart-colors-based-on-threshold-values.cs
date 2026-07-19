// Title: Conditional Formatting Data Bars for a Progress Bar Chart in Aspose.Cells (.NET C#)
// Description: Creates a workbook, writes progress percentages to column A, defines the range A1:A7, and applies three DataBar conditional‑formatting rules that color the bars green (0‑33), yellow (34‑66) and red (67‑100). The file is saved as an XLSX workbook.
// Keywords: Aspose.Cells C# conditional formatting | data bar thresholds | progress bar colors Excel | green yellow red data bars | programmatic Excel styling | .NET Excel API | dynamic progress visualization | Excel conditional formatting API
// Common Searches: Aspose.Cells add data bar conditional formatting C# | how to color progress bars by value in Aspose.Cells | set green yellow red thresholds for Excel data bars | C# code for conditional formatting progress chart | Aspose.Cells conditional formatting multiple rules
// Developer Intent: Generate an Excel workbook that shows progress values with color‑coded data‑bar bars based on numeric thresholds.
// Use Cases: Project status reports that highlight task completion with green, yellow, and red bars. | KPI dashboards where values are instantly recognizable by colored progress bars. | Automated Excel exports that visually encode risk levels using threshold‑based colors.
// AI Prompts: Add a fourth data‑bar rule for values above 100 with a custom purple color. | Show how to hide the numeric value while keeping only the colored bar visible. | Provide code to apply the same three‑rule conditional formatting to multiple worksheets in one workbook. | Explain how to adjust the bar length to use a percentage of the cell width instead of the default setting.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsProgressBarConditionalFormatting
{
    // Creates a workbook, writes progress percentages to column A, defines the range A1:A7, and applies three DataBar conditional‑formatting rules that color the bars green (0‑33), yellow (34‑66) and red (67‑100). The file is saved as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample progress values (0-100)
            double[] progressValues = { 10, 25, 40, 55, 70, 85, 100 };
            for (int i = 0; i < progressValues.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(progressValues[i]);
            }

            // Define the range that will receive the conditional formatting (A1:A7)
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = progressValues.Length - 1,
                StartColumn = 0,
                EndColumn = 0
            };

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(range);

            // ---------- Rule 1: 0 - 33  => Green ----------
            int greenRule = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition greenCondition = cfCollection[greenRule];
            greenCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            greenCondition.DataBar.MinCfvo.Value = 0;
            greenCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            greenCondition.DataBar.MaxCfvo.Value = 33;
            greenCondition.DataBar.Color = Color.Green;
            greenCondition.DataBar.ShowValue = true;

            // ---------- Rule 2: 34 - 66 => Yellow ----------
            int yellowRule = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition yellowCondition = cfCollection[yellowRule];
            yellowCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            yellowCondition.DataBar.MinCfvo.Value = 34;
            yellowCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            yellowCondition.DataBar.MaxCfvo.Value = 66;
            yellowCondition.DataBar.Color = Color.Yellow;
            yellowCondition.DataBar.ShowValue = true;

            // ---------- Rule 3: 67 - 100 => Red ----------
            int redRule = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition redCondition = cfCollection[redRule];
            redCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
            redCondition.DataBar.MinCfvo.Value = 67;
            redCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            redCondition.DataBar.MaxCfvo.Value = 100;
            redCondition.DataBar.Color = Color.Red;
            redCondition.DataBar.ShowValue = true;

            // Save the workbook
            workbook.Save("ProgressBarConditionalFormatting.xlsx", SaveFormat.Xlsx);
        }
    }
}
