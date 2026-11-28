// Title: Generate minimal HTML from an Excel workbook by disabling comments, conditional formatting, and grid lines with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to HTML without comments and gridlines | How to omit conditional formatting when saving Excel as HTML using Aspose.Cells | Create lightweight HTML output from Excel by disabling comments, grid lines, and conditional formatting in Aspose.Cells | Save only the active worksheet as HTML with Aspose.Cells for .NET | Aspose.Cells HtmlSaveOptions minimal HTML export example
// Tags: html export disable comments Aspose.Cells | remove conditional formatting htmlsaveoptions Aspose.Cells | export gridlines false Aspose.Cells HTML | export active worksheet only html Aspose.Cells | minimal html output Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace MinimalHtmlExport
{
    // The example creates a workbook, adds sample data, a comment, and a conditional formatting rule, then uses HtmlSaveOptions to set IsExportComments and ExportGridLines to false, export only the active worksheet, and save the workbook as a minimal HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.25);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.10);

                // Add a comment (will be omitted in HTML)
                int commentIndex = sheet.Comments.Add("A2");
                sheet.Comments[commentIndex].Note = "This is a comment that should be omitted in the HTML output.";

                // Apply a conditional formatting rule (will also be omitted)
                int cfIndex = sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[cfIndex];
                var area = new CellArea { StartRow = 1, StartColumn = 1, EndRow = 3, EndColumn = 1 };
                cf.AddArea(area);
                int conditionIdx = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "1", null);
                cf[conditionIdx].Style.BackgroundColor = System.Drawing.Color.Yellow;

                // Configure HTML save options for minimal output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do not export comments or grid lines
                    IsExportComments = false,
                    ExportGridLines = false,

                    // Export only the active worksheet
                    ExportActiveWorksheetOnly = true,

                    // Set encoding explicitly
                    Encoding = System.Text.Encoding.UTF8
                };

                // Define output file path
                string outputPath = "Minimal.html";

                // Save the workbook as HTML with the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file generated successfully at: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during HTML export:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
