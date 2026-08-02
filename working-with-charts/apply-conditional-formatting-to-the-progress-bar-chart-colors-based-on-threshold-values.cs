// Title: C# – Aspose.Cells for .NET: Apply Threshold‑Based DataBar Conditional Formatting to Create a Progress Bar
// Description: This example creates a new workbook, writes progress percentages to column A, defines the range A1:A6, and adds three DataBar conditional formats—green (0‑50), yellow (51‑80) and red (81‑100). Each cell shows a colored bar with its numeric value, then the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | data bar | progress bar visualization | threshold colors | Excel sample code | GitHub example | API usage
// Common Searches: Aspose.Cells add multiple data bar conditional formats C# | how to color code progress values in Excel with Aspose.Cells | threshold based conditional formatting Aspose.Cells .NET | C# create progress bar chart using data bars | Aspose.Cells conditional formatting range example
// Developer Intent: Generate an Excel file where progress percentages are displayed as colored data bars that change color according to low, medium, and high thresholds.
// Use Cases: Project trackers that highlight task completion with green, yellow, and red bars. | KPI dashboards where metrics are instantly readable through color‑coded data bars. | Printable risk‑assessment reports that emphasize scores using threshold‑based bar colors.
// AI Prompts: Write C# code with Aspose.Cells to add a fourth data‑bar format for values above 100 using a purple color. | Show how to change the conditional‑formatting range from column A to B2:B10 while preserving the existing threshold colors. | Explain how to hide the numeric value and display only the colored data bar in the sample.

using System;
using System.Drawing;
using Aspose.Cells;

// This example creates a new workbook, writes progress percentages to column A, defines the range A1:A6, and adds three DataBar conditional formats—green (0‑50), yellow (51‑80) and red (81‑100). Each cell shows a colored bar with its numeric value, then the workbook is saved as an XLSX file.
class ProgressBarConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample progress values (0‑100) placed in column A
        double[] progressValues = { 20, 45, 60, 75, 90, 30 };
        for (int i = 0; i < progressValues.Length; i++)
        {
            worksheet.Cells[i, 0].PutValue(progressValues[i]);
        }

        // Define the range that will receive the conditional formatting (A1:A6)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = progressValues.Length - 1,
            StartColumn = 0,
            EndColumn = 0
        };

        // Local helper to add a DataBar conditional format with specific min, max and color
        void AddDataBar(double min, double max, Color barColor)
        {
            // Add a new conditional formatting collection
            int fmtIdx = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[fmtIdx];

            // Apply the range to this collection
            fcs.AddArea(range);

            // Add a DataBar condition
            int condIdx = fcs.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcs[condIdx];

            // Configure the DataBar
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.Number;
            dataBar.MinCfvo.Value = min;
            dataBar.MaxCfvo.Type = FormatConditionValueType.Number;
            dataBar.MaxCfvo.Value = max;
            dataBar.Color = barColor;
            dataBar.ShowValue = true; // display the numeric value alongside the bar
        }

        // Low progress (0‑50) → Green bar
        AddDataBar(0, 50, Color.Green);

        // Medium progress (51‑80) → Yellow bar
        AddDataBar(51, 80, Color.Yellow);

        // High progress (81‑100) → Red bar
        AddDataBar(81, 100, Color.Red);

        // Save the workbook
        workbook.Save("ProgressBarConditionalFormatting.xlsx", SaveFormat.Xlsx);
    }
}
