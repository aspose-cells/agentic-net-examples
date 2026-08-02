// Title: Apply Conditional Formatting and Export to HTML with Aspose.Cells for .NET
// Description: Demonstrates how to add a conditional‑formatting rule (value > 20) to a range, configure HtmlSaveOptions (MergeAreas, ExportGridLines, CSS), and save the workbook as an HTML file while preserving the visual styling.
// Keywords: Aspose.Cells C# conditional formatting HTML export | HtmlSaveOptions MergeAreas | ExportGridLines Aspose.Cells | preserve Excel formatting in HTML | conditional formatting to HTML .NET | Aspose.Cells HTMLSaveOptions CSS | export active worksheet HTML
// Common Searches: Aspose.Cells keep conditional formatting when converting to HTML | HtmlSaveOptions MergeAreas example | Export grid lines with Aspose.Cells HTML | C# export worksheet with conditional formatting to HTML | How to use CSS instead of inline styles in Aspose.Cells HTML output
// Developer Intent: The developer wants to apply a conditional‑formatting rule to a cell range and generate an HTML file that displays the same visual highlights as the Excel workbook.
// Use Cases: Create an HTML report where cells above a threshold are highlighted with custom colors and bold text. | Export only the active worksheet while preserving conditional formatting for web‑based dashboards. | Generate clean HTML using CSS (not inline styles) and retain grid lines for a spreadsheet‑like appearance.
// AI Prompts: Show C# code that adds a conditional formatting rule for values greater than 20 and saves the workbook to HTML with formatting preserved using Aspose.Cells. | Provide an example of HtmlSaveOptions settings (MergeAreas, ExportGridLines, DisableCss) for exporting a worksheet with conditional formatting. | Explain how the MergeAreas property influences the HTML output when conditional formatting is applied.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to add a conditional‑formatting rule (value > 20) to a range, configure HtmlSaveOptions (MergeAreas, ExportGridLines, CSS), and save the workbook as an HTML file while preserving the visual styling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(25);
            worksheet.Cells["A3"].PutValue(15);
            worksheet.Cells["B1"].PutValue(5);
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(12);

            // ---------- Apply Conditional Formatting ----------
            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var conditionalFormatting = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A1:B3)
            CellArea area = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                StartColumn = 0, // Column A
                EndRow = 2,     // Row 3
                EndColumn = 1   // Column B
            };
            conditionalFormatting.AddArea(area);

            // Create a condition: cell value greater than 20
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "20",
                null);

            // Define the style that will be applied when the condition is met
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.LightCoral;   // background color
            highlightStyle.Pattern = BackgroundType.Solid;      // solid fill
            highlightStyle.Font.Color = Color.White;            // font color
            highlightStyle.Font.IsBold = true;

            // Attach the style to the conditional formatting rule
            FormatCondition condition = conditionalFormatting[conditionIndex];
            condition.Style = highlightStyle;
            // ----------------------------------------------------

            // ---------- Configure HTML save options ----------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Merge conditional formatting areas before saving to preserve visual appearance
                MergeAreas = true,

                // Export grid lines so the HTML layout matches Excel more closely
                ExportGridLines = true,

                // Export only the active worksheet (optional, can be set to false to export all)
                ExportActiveWorksheetOnly = true,

                // Use CSS (instead of inline styles) for cleaner HTML; set to true if you prefer inline
                DisableCss = false
            };
            // --------------------------------------------------

            // Save the workbook as an HTML file with the configured options
            workbook.Save("ConditionalFormatting.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
