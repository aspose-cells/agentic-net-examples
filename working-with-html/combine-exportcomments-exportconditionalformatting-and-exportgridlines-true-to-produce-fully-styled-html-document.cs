// Title: Export Fully Styled HTML (comments, conditional formatting, gridlines) with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, add sample data, a cell comment, a red‑background conditional format for values > 50, enable gridlines, and save the worksheet as a single‑page HTML file that preserves all styles using HtmlSaveOptions (IsExportComments, ExportGridLines, ExportDataOptions.All).
// Keywords: Aspose.Cells | C# | HTML export | export comments | conditional formatting | gridlines | HtmlSaveOptions | ExportDataOptions.All | .NET | Excel to HTML | full style export
// Common Searches: Aspose.Cells export comments to HTML C# | How to include conditional formatting in HTML export Aspose.Cells | Export Excel gridlines to HTML using Aspose.Cells .NET | Save workbook as styled HTML Aspose.Cells | HtmlSaveOptions IsExportComments example
// Developer Intent: Create an HTML version of an Excel worksheet that retains cell comments, conditional formatting rules, and visible gridlines.
// Use Cases: Web dashboards that require exact Excel styling with notes and color cues. | Automated email reports where comments and conditional highlights must remain visible. | Documentation pages that showcase threshold‑based coloring directly from the source workbook. | Embedding Excel data in web applications without losing formatting.
// AI Prompts: Generate C# code using Aspose.Cells to export a worksheet to HTML with comments, conditional formatting, and gridlines preserved. | Explain each HtmlSaveOptions property (IsExportComments, ExportGridLines, ExportDataOptions) and its effect on the output HTML. | Show how to add a comment and a >50 conditional formatting rule before saving as a fully styled HTML document. | Provide a step‑by‑step guide to enable gridlines visibility and export them in HTML with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook, add sample data, a cell comment, a red‑background conditional format for values > 50, enable gridlines, and save the worksheet as a single‑page HTML file that preserves all styles using HtmlSaveOptions (IsExportComments, ExportGridLines, ExportDataOptions.All).
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(70);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(45);

                // Add a comment to cell A2
                int commentIndex = sheet.Comments.Add("A2");
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = "Seasonal fruit";

                // Apply conditional formatting: highlight quantities > 50 with red background
                int startRow = 1;          // zero‑based index (row 2 in Excel)
                int startColumn = 1;       // column B
                int totalRows = 3;         // rows 2‑4
                int totalColumns = 1;      // only column B

                // Define the range for conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = startRow,
                    StartColumn = startColumn,
                    EndRow = startRow + totalRows - 1,
                    EndColumn = startColumn + totalColumns - 1
                };

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Associate the defined range with the conditional formatting
                fcc.AddArea(area);

                // Add a condition: cell value greater than 50
                int conditionIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition condition = fcc[conditionIndex];

                // Define the style for the condition (red background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Ensure gridlines are visible in the worksheet (optional, but we also export them)
                sheet.IsGridlinesVisible = true;

                // Configure HTML save options to export comments, gridlines, and all data
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    IsExportComments = true,                 // Export cell comments
                    ExportGridLines = true,                  // Export worksheet gridlines
                    ExportActiveWorksheetOnly = true,        // Export only this worksheet
                    ExportDataOptions = HtmlExportDataOptions.All // Export all data (including styles)
                };

                // Save the workbook as an HTML file with the specified options
                workbook.Save("FullyStyledOutput.html", htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
