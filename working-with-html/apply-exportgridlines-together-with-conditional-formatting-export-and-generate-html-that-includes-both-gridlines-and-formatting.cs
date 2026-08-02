// Title: Export Excel to HTML with Gridlines & Conditional Formatting – Aspose.Cells C# Example
// Description: Demonstrates how to create a workbook, enable gridlines, apply a conditional formatting rule (value > 20), and save it as HTML using Aspose.Cells for .NET. HtmlSaveOptions.ExportGridLines is set to true and ExportDataOptions to All, ensuring both gridlines and formatting are retained.
// Keywords: Aspose.Cells | C# HTML export | ExportGridLines | conditional formatting | HtmlSaveOptions | Excel to HTML | gridlines visible | HtmlExportDataOptions.All | .NET spreadsheet conversion | web report generation
// Common Searches: Aspose.Cells export gridlines to HTML | C# export Excel with conditional formatting to HTML | HtmlSaveOptions ExportGridLines example | How to keep Excel gridlines in HTML output | Conditional formatting not showing in Aspose.Cells HTML export
// Developer Intent: Create an HTML representation of an Excel worksheet that preserves the original gridlines and conditional formatting using Aspose.Cells for .NET.
// Use Cases: Web dashboards that need exact Excel layout with gridlines and highlighted cells | Automated email reports where Excel sheets are converted to HTML while keeping visual cues | Documentation snapshots of spreadsheets for intranet publishing with full formatting
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to HTML with ExportGridLines enabled and conditional formatting applied. | Explain the role of HtmlSaveOptions.ExportGridLines and HtmlExportDataOptions.All in preserving worksheet appearance. | Troubleshoot why conditional formatting might be missing in the exported HTML and suggest fixes.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, enable gridlines, apply a conditional formatting rule (value > 20), and save it as HTML using Aspose.Cells for .NET. HtmlSaveOptions.ExportGridLines is set to true and ExportDataOptions to All, ensuring both gridlines and formatting are retained.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Make gridlines visible in the worksheet
                sheet.IsGridlinesVisible = true;

                // Populate some sample data
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue(row * col);
                    }
                }

                // Define the range for conditional formatting (A1:E10)
                int cfIndex = sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[cfIndex];
                cf.AddArea(new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 4 });

                // Add a condition: highlight cells with value > 20
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "20",
                    string.Empty); // formula2 not used for this operator

                // Retrieve the created condition
                FormatCondition condition = cf[conditionIndex];

                // Define the style for the condition
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightCoral;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Configure HTML save options to export gridlines and all data (including conditional formatting)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,
                    ExportDataOptions = HtmlExportDataOptions.All
                };

                // Save the workbook as an HTML file
                string outputPath = "ConditionalFormattingWithGridlines.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
