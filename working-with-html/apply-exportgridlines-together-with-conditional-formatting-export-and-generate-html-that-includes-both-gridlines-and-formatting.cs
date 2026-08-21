// Title: Export Gridlines and Conditional Formatting to HTML with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, enable worksheet gridlines, apply a conditional formatting rule (values > 20 highlighted in orange), and save the sheet as an HTML file using Aspose.Cells HtmlSaveOptions with ExportGridLines, row/column headings, and active‑worksheet only settings.
// Keywords: Aspose.Cells HTML export | ExportGridLines C# | conditional formatting to HTML | HtmlSaveOptions example | gridlines in HTML output | .NET Excel to HTML | Aspose.Cells workbook to HTML | C# Excel export with styling
// Common Searches: Aspose.Cells export gridlines to HTML | how to keep conditional formatting when saving as HTML | C# HtmlSaveOptions ExportGridLines sample | export Excel worksheet with row and column headings HTML | Aspose.Cells HTML export with styling
// Developer Intent: Generate an HTML representation of an Excel worksheet that shows both gridlines and conditional formatting.
// Use Cases: Web dashboards that need Excel‑style gridlines and highlighted cells. | Embedding styled Excel data in emails or web pages without losing formatting. | Creating printable HTML snapshots of reports with value‑based highlights.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to HTML with gridlines and conditional formatting preserved. | Show how to add a conditional formatting rule for values greater than a threshold and save the sheet as HTML with row/column headings. | Explain the HtmlSaveOptions settings required to include gridlines, export only the active worksheet, and retain conditional formatting.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, enable worksheet gridlines, apply a conditional formatting rule (values > 20 highlighted in orange), and save the sheet as an HTML file using Aspose.Cells HtmlSaveOptions with ExportGridLines, row/column headings, and active‑worksheet only settings.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (values 1 to 100)
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        worksheet.Cells[row, col].PutValue(row * 3 + col + 1);
                    }
                }

                // Enable gridlines visibility in the worksheet
                worksheet.IsGridlinesVisible = true;

                // Add a simple conditional formatting rule:
                // Highlight cells with value greater than 20 with a light orange background
                int cfIndex = worksheet.ConditionalFormattings.Add();
                var cf = worksheet.ConditionalFormattings[cfIndex];

                // Apply to the populated range A1:C10
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 2
                };
                cf.AddArea(area);

                // Add condition (operator GreaterThan requires only one formula; second can be null)
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "20",
                    null);
                FormatCondition condition = cf[conditionIndex];

                // Define the style for the condition
                Style cfStyle = workbook.CreateStyle();
                cfStyle.ForegroundColor = Color.FromArgb(255, 230, 180); // Light orange
                cfStyle.Pattern = BackgroundType.Solid;
                condition.Style = cfStyle;

                // Configure HTML save options to export gridlines
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,               // Export the worksheet gridlines
                    ExportActiveWorksheetOnly = true,     // Export only the first worksheet
                    ExportRowColumnHeadings = true        // Include row/column headings
                };

                // Save the workbook as HTML with the specified options
                string outputPath = "ConditionalFormattingWithGridlines.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file generated at '{outputPath}' with gridlines and conditional formatting.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
