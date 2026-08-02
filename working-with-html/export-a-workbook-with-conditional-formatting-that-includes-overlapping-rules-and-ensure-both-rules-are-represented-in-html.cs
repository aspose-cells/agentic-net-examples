// Title: Export Overlapping Conditional Formatting to HTML with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, apply two overlapping conditional‑formatting rules (value > 50 → red, value > 30 → green) with StopIfTrue disabled, configure HtmlSaveOptions to retain all styles and export worksheet CSS separately, and save the result as an HTML file that displays both formats.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | overlapping rules | HTML export | HtmlSaveOptions | ExcludeUnusedStyles | ExportWorksheetCSSSeparately | StopIfTrue | Excel to web | workbook to HTML
// Common Searches: Aspose.Cells preserve overlapping conditional formats in HTML | HtmlSaveOptions ExcludeUnusedStyles false example | Export worksheet CSS separately Aspose.Cells .NET | How to keep multiple conditional formatting rules when saving as HTML | C# export Excel conditional formatting to HTML with Aspose
// Developer Intent: Generate an HTML representation of an Excel workbook that shows all overlapping conditional‑formatting rules applied to the same cell range.
// Use Cases: Web dashboards that need layered color thresholds for the same data range. | Auditable financial reports where every conditional rule must be visible in the HTML view. | Documentation portals that render Excel sheets as HTML while preserving complex formatting logic.
// AI Prompts: Add a third conditional formatting rule (e.g., value > 10 → yellow) and keep all three rules visible in the exported HTML. | Show how to embed the generated CSS inline instead of separate files while still displaying overlapping formats. | Explain the effect of the StopIfTrue property on rule evaluation order during HTML export with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply two overlapping conditional‑formatting rules (value > 50 → red, value > 30 → green) with StopIfTrue disabled, configure HtmlSaveOptions to retain all styles and export worksheet CSS separately, and save the result as an HTML file that displays both formats.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill cells A1:A5 with sample numeric values
        // Values: 10, 30, 50, 70, 90
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 20 + 10);
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A5 for both rules
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 4,
            StartColumn = 0,
            EndColumn = 0
        };
        cfCollection.AddArea(range);

        // -------------------------------------------------
        // First rule: cells with value > 50 get a red background
        // -------------------------------------------------
        int condIndex1 = cfCollection.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "50",
            null);
        FormatCondition condition1 = cfCollection[condIndex1];
        condition1.Style.BackgroundColor = Color.Red;
        // Allow lower‑priority rules to be evaluated as well
        condition1.StopIfTrue = false;

        // -------------------------------------------------
        // Second rule: cells with value > 30 get a green background
        // This rule overlaps with the first one (e.g., value 70 satisfies both)
        // -------------------------------------------------
        int condIndex2 = cfCollection.AddCondition(
            FormatConditionType.CellValue,
            OperatorType.GreaterThan,
            "30",
            null);
        FormatCondition condition2 = cfCollection[condIndex2];
        condition2.Style.BackgroundColor = Color.Green;
        condition2.StopIfTrue = false;

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Keep all generated styles so both conditional formats appear in the HTML
        htmlOptions.ExcludeUnusedStyles = false;
        // Export worksheet CSS separately for clearer inspection (optional)
        htmlOptions.ExportWorksheetCSSSeparately = true;

        // Save the workbook as HTML
        workbook.Save("ConditionalFormattingOverlap.html", htmlOptions);
    }
}
