// Title: Export an Excel workbook to HTML with bold and italic text preserved via CSS using Aspose.Cells for .NET
// Common Searches: how to save an Aspose.Cells workbook as HTML with bold and italic formatting in C# | Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately example C# | C# export Excel to HTML retaining font styles using CSS classes Aspose.Cells | generate HTML from workbook with separate CSS file and Base64 images using Aspose.Cells | preserve cell font bold/italic when converting Excel to HTML with Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions CSS output | Excel to HTML conversion with bold/italic retention | ExportWorksheetCSSSeparately usage example | Base64 image embedding in HTML export | grid line visibility in HTML workbook export

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The sample creates a workbook, applies bold, italic, and combined bold‑italic styles to cells, and saves the workbook as HTML using HtmlSaveOptions configured to export worksheet CSS separately, keep grid lines visible, and embed images as Base64, ensuring the font styles are retained through generated CSS classes.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with sample text
                sheet.Cells["A1"].PutValue("Normal Text");
                sheet.Cells["A2"].PutValue("Bold Text");
                sheet.Cells["A3"].PutValue("Italic Text");
                sheet.Cells["A4"].PutValue("Bold & Italic Text");

                // Apply bold style to A2
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;
                sheet.Cells["A2"].SetStyle(boldStyle);

                // Apply italic style to A3
                Style italicStyle = workbook.CreateStyle();
                italicStyle.Font.IsItalic = true;
                sheet.Cells["A3"].SetStyle(italicStyle);

                // Apply both bold and italic to A4
                Style boldItalicStyle = workbook.CreateStyle();
                boldItalicStyle.Font.IsBold = true;
                boldItalicStyle.Font.IsItalic = true;
                sheet.Cells["A4"].SetStyle(boldItalicStyle);

                // Configure HTML save options to retain font styles via CSS
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export the worksheet CSS into a separate file (demonstrates CSS usage)
                    ExportWorksheetCSSSeparately = true,

                    // Keep grid lines visible for easier visual verification
                    ExportGridLines = true,

                    // Export images as Base64 to keep the HTML self‑contained
                    ExportImagesAsBase64 = true
                };

                // Define output path for the HTML file (CSS will be generated automatically)
                string outputHtml = "ExportedWorkbook.html";

                // Save the workbook as HTML
                workbook.Save(outputHtml, htmlOptions);

                Console.WriteLine($"HTML file saved to: {outputHtml}");
                Console.WriteLine("Bold and italic formatting are retained via CSS classes.");
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
