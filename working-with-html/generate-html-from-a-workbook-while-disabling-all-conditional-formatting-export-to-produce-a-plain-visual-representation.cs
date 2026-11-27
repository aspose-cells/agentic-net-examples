// Title: Generate HTML from an Aspose.Cells workbook while disabling conditional formatting in C#
// Common Searches: Aspose.Cells C# export workbook to HTML without applying conditional formatting | How to clear all conditional formatting from a worksheet before saving as HTML using Aspose.Cells | C# code to remove conditional formatting collection and produce plain HTML output | Export Excel to HTML with grid lines and base64 images while ignoring conditional formats in Aspose.Cells | Aspose.Cells HtmlSaveOptions example for plain visual HTML representation of a workbook
// Tags: Aspose.Cells HTML export without conditional formatting | HtmlSaveOptions embed images base64 C# | remove worksheet conditional formatting Aspose.Cells | export grid lines in HTML Aspose.Cells | plain HTML view of Excel workbook C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // The sample creates a workbook, adds data and a conditional formatting rule, saves the original file as XLSX, clears all conditional formatting from the worksheet, configures HtmlSaveOptions to export grid lines and embed images as base64, and finally saves the workbook as an HTML file that shows a plain visual representation without any conditional formatting.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define paths for the output files
                string outputHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "WorkbookWithoutConditionalFormatting.html");
                string outputXlsxPath = Path.Combine(Directory.GetCurrentDirectory(), "WorkbookWithConditionalFormatting.xlsx");

                // ------------------------------------------------------------
                // 1. Create a sample workbook and add some data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (A1:B6)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                cells["A2"].PutValue("Apples");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Bananas");
                cells["B3"].PutValue(55);
                cells["A4"].PutValue("Cherries");
                cells["B4"].PutValue(70);
                cells["A5"].PutValue("Dates");
                cells["B5"].PutValue(20);
                cells["A6"].PutValue("Elderberries");
                cells["B6"].PutValue(90);

                // ------------------------------------------------------------
                // 2. Add a simple conditional formatting rule:
                //    Highlight cells in column B with a value > 50 (red fill)
                // ------------------------------------------------------------
                int cfIndex = sheet.ConditionalFormattings.Add();
                // The collection element is a FormatConditionCollection
                FormatConditionCollection cf = sheet.ConditionalFormattings[cfIndex];

                // Define the range to which the rule applies (B2:B6)
                CellArea area = new CellArea
                {
                    StartRow = 1,
                    EndRow = 5,
                    StartColumn = 1,
                    EndColumn = 1
                };
                cf.AddArea(area);

                // Add a condition: Cell value > 50
                int conditionIdx = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition condition = cf[conditionIdx];

                // Apply a red background style
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Save the workbook with conditional formatting (for reference)
                workbook.Save(outputXlsxPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook with conditional formatting saved to: {outputXlsxPath}");

                // ------------------------------------------------------------
                // 3. Remove conditional formatting before exporting to HTML
                // ------------------------------------------------------------
                sheet.ConditionalFormattings.Clear();

                // ------------------------------------------------------------
                // 4. Prepare HTML save options
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Keep grid lines visible for better readability
                    ExportGridLines = true,
                    // Embed images as base64 to keep a single HTML file
                    ExportImagesAsBase64 = true
                };

                // ------------------------------------------------------------
                // 5. Save the workbook as HTML (plain visual representation)
                // ------------------------------------------------------------
                workbook.Save(outputHtmlPath, htmlOptions);
                Console.WriteLine($"HTML file saved without conditional formatting to: {outputHtmlPath}");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors to the console
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
