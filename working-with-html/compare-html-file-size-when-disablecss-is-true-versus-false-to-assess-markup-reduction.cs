// Title: Aspose.Cells C# – Compare HTML output size with HtmlSaveOptions.DisableCss true vs false
// Description: This example creates a workbook with styled cells, saves it twice as HTML—once with inline styles (DisableCss=true) and once with a CSS block (DisableCss=false)—then reads the file sizes and reports the byte and percentage reduction, demonstrating how the DisableCss option influences markup size.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | HTML export size | inline styles vs CSS | markup reduction | C# Excel to HTML | file size comparison
// Common Searches: Aspose.Cells compare HTML size DisableCss true false | HTML export inline styles size Aspose.Cells | Effect of DisableCss on generated HTML markup | Measure HTML file size reduction Aspose.Cells | C# Aspose.Cells HTML size optimization
// Developer Intent: Evaluate the impact of the HtmlSaveOptions.DisableCss setting on the size of HTML generated from a workbook.
// Use Cases: Choose the most compact HTML format for email or web embedding. | Automate a regression test that flags unexpected markup growth. | Guide performance tuning by quantifying the trade‑off between inline styles and external CSS.
// AI Prompts: Generate C# code that calculates and displays the percentage reduction between the two HTML files. | Explain how Aspose.Cells builds CSS when DisableCss is false and why it affects file size. | Recommend additional HtmlSaveOptions (e.g., ExportImagesAsBase64, Minify) to further shrink the HTML output while keeping visual fidelity.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook with styled cells, saves it twice as HTML—once with inline styles (DisableCss=true) and once with a CSS block (DisableCss=false)—then reads the file sizes and reports the byte and percentage reduction, demonstrating how the DisableCss option influences markup size.
    public class HtmlDisableCssComparison
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some styled data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with various formats to generate noticeable HTML markup
                sheet.Cells["A1"].PutValue("Bold Red Text");
                var styleA1 = sheet.Cells["A1"].GetStyle();
                styleA1.Font.IsBold = true;
                styleA1.Font.Color = System.Drawing.Color.Red;
                sheet.Cells["A1"].SetStyle(styleA1);

                sheet.Cells["B2"].PutValue("Italic Blue Text");
                var styleB2 = sheet.Cells["B2"].GetStyle();
                styleB2.Font.IsItalic = true;
                styleB2.Font.Color = System.Drawing.Color.Blue;
                sheet.Cells["B2"].SetStyle(styleB2);

                sheet.Cells["C3"].PutValue("Large Green Text");
                var styleC3 = sheet.Cells["C3"].GetStyle();
                styleC3.Font.Size = 16;
                styleC3.Font.Color = System.Drawing.Color.Green;
                sheet.Cells["C3"].SetStyle(styleC3);

                // Prepare HTML save options with DisableCss = true (inline styles only)
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    DisableCss = true
                };
                string inlineHtmlPath = "HtmlWithInlineStyles.html";
                workbook.Save(inlineHtmlPath, options);

                // Change option to use CSS (DisableCss = false) and save again
                options.DisableCss = false;
                string cssHtmlPath = "HtmlWithCssStyles.html";
                workbook.Save(cssHtmlPath, options);

                // Get file sizes
                long inlineSize = new FileInfo(inlineHtmlPath).Length;
                long cssSize = new FileInfo(cssHtmlPath).Length;

                // Output comparison results
                Console.WriteLine($"File size with inline styles (DisableCss=true): {inlineSize} bytes");
                Console.WriteLine($"File size with CSS styles (DisableCss=false): {cssSize} bytes");
                Console.WriteLine($"Size reduction: {cssSize - inlineSize} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HtmlDisableCssComparison.Run();
        }
    }
}
