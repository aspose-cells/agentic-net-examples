// Title: Apply threshold‑based red, yellow, and green data‑bar conditional formatting in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes progress percentages to column A, defines the target range, and adds three DataBar conditional‑format rules (0‑30 → red, 31‑70 → yellow, 71‑100 → green). The formatted column is saved as an XLSX file.
// Keywords: Aspose.Cells C# conditional formatting | data bar thresholds Aspose.Cells | progress bar colors Excel .NET | multi‑range data bar Aspose | red yellow green conditional format | Aspose.Cells progress visualization
// Common Searches: Aspose.Cells data bar conditional formatting C# | how to set red yellow green progress bars in Excel using Aspose | multiple data bar rules Aspose.Cells .NET | conditional formatting thresholds Aspose.Cells example
// Developer Intent: Generate an Excel file with a progress column where values are automatically colored red, yellow, or green via data‑bar conditional formatting based on defined numeric ranges.
// Use Cases: Project dashboards that highlight task completion status with color‑coded bars. | KPI reports where low, medium, and high metrics are instantly recognizable. | Sales or performance sheets that flag under‑performing, average, and top results.
// AI Prompts: Write C# code using Aspose.Cells to add three DataBar conditional formats with custom colors for low, medium, and high ranges. | Show how to extend the example with an additional rule for values above 100 (e.g., blue color). | Provide a guide to list, modify, or delete existing DataBar conditional formats in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, writes progress percentages to column A, defines the target range, and adds three DataBar conditional‑format rules (0‑30 → red, 31‑70 → yellow, 71‑100 → green). The formatted column is saved as an XLSX file.
class ProgressBarConditionalFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample progress values (0-100)
        double[] progressValues = { 10, 35, 55, 80, 20, 65, 90 };
        for (int i = 0; i < progressValues.Length; i++)
        {
            sheet.Cells[i, 0].PutValue(progressValues[i]);
        }

        // Define the range that will receive the data‑bar conditional formatting
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

        // ---------- Low range (0 – 30) : Red ----------
        int lowIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
        FormatCondition lowCondition = cfCollection[lowIdx];
        lowCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
        lowCondition.DataBar.MinCfvo.Value = 0;
        lowCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
        lowCondition.DataBar.MaxCfvo.Value = 30;
        lowCondition.DataBar.Color = Color.Red;
        lowCondition.DataBar.ShowValue = true; // show the numeric value

        // ---------- Medium range (31 – 70) : Yellow ----------
        int mediumIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
        FormatCondition mediumCondition = cfCollection[mediumIdx];
        mediumCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
        mediumCondition.DataBar.MinCfvo.Value = 31;
        mediumCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
        mediumCondition.DataBar.MaxCfvo.Value = 70;
        mediumCondition.DataBar.Color = Color.Yellow;
        mediumCondition.DataBar.ShowValue = true;

        // ---------- High range (71 – 100) : Green ----------
        int highIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
        FormatCondition highCondition = cfCollection[highIdx];
        highCondition.DataBar.MinCfvo.Type = FormatConditionValueType.Number;
        highCondition.DataBar.MinCfvo.Value = 71;
        highCondition.DataBar.MaxCfvo.Type = FormatConditionValueType.Number;
        highCondition.DataBar.MaxCfvo.Value = 100;
        highCondition.DataBar.Color = Color.Green;
        highCondition.DataBar.ShowValue = true;

        // Save the workbook
        workbook.Save("ProgressBarConditionalFormatting.xlsx", SaveFormat.Xlsx);
    }
}
