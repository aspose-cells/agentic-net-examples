// Title: How to export Excel grid lines and conditional formatting to HTML with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to HTML with grid lines and conditional formatting | Save Excel file as HTML preserving grid lines using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportGridLines with conditional formatting example | Generate HTML from Excel that shows cell borders and conditional colors in C# | Export Excel conditional formatting to HTML while keeping grid lines using Aspose.Cells
// Tags: ExportGridLines HtmlSaveOptions | conditional formatting HTML export | grid lines in HTML output | Aspose.Cells HTML conversion with styles | presentation preference option

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace ExportGridLinesWithConditionalFormatting
{
    // The example creates a workbook, fills cells A1:C10 with numeric data, adds a conditional formatting rule that highlights values greater than 30 with a light‑salmon background, configures HtmlSaveOptions to export grid lines, embed images as Base64, export only the active worksheet, set a page title, and enable presentation preference, then saves the workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Output HTML file path
            string outputHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalFormattingWithGridLines.html");

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(outputHtmlPath));

                // ------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill numeric data in range A1:C10
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        cells[row, col].PutValue(row * 10 + col * 5);
                    }
                }

                // ------------------------------------------------------------
                // 2. Add a conditional formatting rule (value > 30, light salmon fill)
                // ------------------------------------------------------------
                // Ensure the ConditionalFormattings collection exists
                sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[0];

                // Apply the rule to the range A1:C10
                cf.AddArea(new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 2 });

                // Add condition: Cell Value > 30
                int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "30", null);
                FormatCondition condition = cf[conditionIndex];

                // Define the style (light salmon background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightSalmon;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // ------------------------------------------------------------
                // 3. Configure HTML save options (export grid lines & conditional formatting)
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,                     // Show grid lines
                    ExportActiveWorksheetOnly = true,           // Export only the active sheet
                    ExportImagesAsBase64 = true,                // Embed images as Base64
                    PageTitle = "Conditional Formatting with Grid Lines",
                    PresentationPreference = true               // Better visual layout
                };

                // ------------------------------------------------------------
                // 4. Save the workbook as HTML
                // ------------------------------------------------------------
                workbook.Save(outputHtmlPath, htmlOptions);

                Console.WriteLine($"HTML file successfully saved to: {outputHtmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting HTML:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
