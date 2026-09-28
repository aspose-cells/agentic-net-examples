// Title: Measure HTML file size difference with and without CSS using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, saves it to HTML twice—once with HtmlSaveOptions.DisableCss set to false and once set to true—and outputs the byte size of each file. | Show how to use MemoryStream to capture the HTML output from Aspose.Cells and compute the reduction in markup size when CSS is disabled.
// Common Searches: c# Aspose.Cells compare HTML export size when DisableCss is true versus false | how to reduce HTML markup size in Aspose.Cells by disabling CSS | measure byte size difference of HTML files generated with Aspose.Cells HtmlSaveOptions.DisableCss
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss | C# HTML export size measurement | Aspose.Cells markup reduction | MemoryStream HTML size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a workbook, saves it to HTML twice—first with CSS included (DisableCss = false) and then without CSS (DisableCss = true)—using MemoryStream, and prints the byte sizes and the reduction achieved.
class HtmlCssComparison
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Row 1");
        sheet.Cells["A3"].PutValue("Row 2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["B3"].PutValue(456);

        // Save to HTML with CSS enabled (default)
        HtmlSaveOptions optionsWithCss = new HtmlSaveOptions();
        optionsWithCss.DisableCss = false; // CSS will be included
        using (MemoryStream msWithCss = new MemoryStream())
        {
            workbook.Save(msWithCss, optionsWithCss);
            long sizeWithCss = msWithCss.Length;

            // Save to HTML with CSS disabled
            HtmlSaveOptions optionsWithoutCss = new HtmlSaveOptions();
            optionsWithoutCss.DisableCss = true; // CSS will be omitted
            using (MemoryStream msWithoutCss = new MemoryStream())
            {
                workbook.Save(msWithoutCss, optionsWithoutCss);
                long sizeWithoutCss = msWithoutCss.Length;

                // Output the file sizes for comparison
                Console.WriteLine($"HTML size with CSS:    {sizeWithCss} bytes");
                Console.WriteLine($"HTML size without CSS: {sizeWithoutCss} bytes");
                Console.WriteLine($"Reduction: {sizeWithCss - sizeWithoutCss} bytes");
            }
        }
    }
}
