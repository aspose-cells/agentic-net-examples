// Title: Apply Conditional Formatting to Highlight High Values and Export Workbook as HTML with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export workbook to HTML while keeping conditional formatting | How to use HtmlSaveOptions to preserve cell styles when converting Excel to HTML in .NET | Aspose.Cells highlight cells greater than 50 before saving as HTML | Export Excel grid lines and embedded images as Base64 using Aspose.Cells HtmlSaveOptions | Create sample workbook and add conditional formatting programmatically with Aspose.Cells C#
// Tags: conditional-formatting html export Aspose.Cells | HtmlSaveOptions embed images base64 C# | gridlines export Aspose.Cells HTML | style preservation Excel to HTML Aspose.Cells | programmatic conditional formatting workbook C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingHtmlExport
{
    // The example loads or creates an Excel workbook, adds conditional formatting that highlights values over 50 in column B, configures HtmlSaveOptions to keep grid lines and embed images as Base64, and saves the workbook as an HTML file while preserving the visual cues.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            try
            {
                // Load or create a workbook
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Create a sample workbook if the input file does not exist
                    workbook = CreateSampleWorkbook();
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook created and saved to \"{inputPath}\".");
                }

                // Apply conditional formatting to preserve visual cues in HTML
                ApplyConditionalFormatting(workbook);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,               // Keep grid lines visible
                    ExportImagesAsBase64 = true,          // Embed images directly in HTML
                    ExportActiveWorksheetOnly = false     // Export the whole workbook
                    // Conditional formatting is exported by default; other properties omitted for compatibility
                };

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["A2"].PutValue("Apples");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Bananas");
            ws.Cells["B3"].PutValue(75);
            ws.Cells["A4"].PutValue("Cherries");
            ws.Cells["B4"].PutValue(120);
            ws.Cells["A5"].PutValue("Dates");
            ws.Cells["B5"].PutValue(45);

            return wb;
        }

        private static void ApplyConditionalFormatting(Workbook workbook)
        {
            try
            {
                Worksheet ws = workbook.Worksheets[0];

                // Add a new ConditionalFormatting object to the worksheet
                int cfIndex = ws.ConditionalFormattings.Add();
                var cf = ws.ConditionalFormattings[cfIndex]; // Use var to avoid explicit type reference

                // Define the range B2:B5 (zero‑based indices)
                CellArea area = new CellArea
                {
                    StartRow = 1,   // B2
                    StartColumn = 1,
                    EndRow = 4,     // B5
                    EndColumn = 1
                };
                cf.AddArea(area);

                // Add a condition: Cell value > 50
                int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition condition = cf[conditionIndex];

                // Define the style to apply when the condition is met
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                Console.WriteLine("Conditional formatting applied: highlight cells in B2:B5 with value > 50.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error applying conditional formatting:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
