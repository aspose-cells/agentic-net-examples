// Title: Aspose.Cells C# – Compare HTML Size with DisableCss True vs False
// Description: C# example that creates a workbook, applies bold, colored, italic and underline formatting, then saves the sheet to HTML twice – once with inline styles (DisableCss = true) and once with external CSS (DisableCss = false). The code reads both file sizes and reports which option yields a smaller markup.
// Keywords: Aspose.Cells HtmlSaveOptions | DisableCss | inline styles | external CSS | HTML file size | markup reduction | C# Excel to HTML | Aspose.Cells performance
// Common Searches: Aspose.Cells DisableCss size comparison | HTML output smaller with inline styles Aspose.Cells | C# compare Excel to HTML file size | Does DisableCss reduce HTML size in Aspose.Cells | Aspose.Cells HtmlSaveOptions CSS vs inline
// Developer Intent: Find out whether disabling CSS (inline styling) or keeping CSS external produces a smaller HTML file when exporting Excel with Aspose.Cells.
// Use Cases: Choose the most bandwidth‑efficient HTML export for email templates generated from Excel workbooks. | Optimize web‑application performance by selecting the smaller HTML representation for Excel‑derived content. | Automate batch processing of workbooks to programmatically decide between inline styles and external CSS based on file‑size thresholds.
// AI Prompts: Generate C# code that calculates the percentage size difference between the inline‑style HTML file and the CSS‑style HTML file produced by Aspose.Cells. | Write a PowerShell script that iterates over all .xlsx files in a directory, saves each as HTML with DisableCss true and false, and outputs a summary table of file sizes. | Explain how to modify the sample to export CSS to a separate .css file, reference it from the HTML, and compare the combined size of HTML + CSS versus the inline‑style HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlSizeComparison
{
    // C# example that creates a workbook, applies bold, colored, italic and underline formatting, then saves the sheet to HTML twice – once with inline styles (DisableCss = true) and once with external CSS (DisableCss = false). The code reads both file sizes and reports which option yields a smaller markup.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data and apply various formatting
            // This ensures that the generated HTML contains both inline styles and CSS rules
            sheet.Cells["A1"].PutValue("Header");
            Style headerStyle = sheet.Cells["A1"].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.Blue;
            headerStyle.Font.Size = 14;
            sheet.Cells["A1"].SetStyle(headerStyle);

            sheet.Cells["B2"].PutValue("Important");
            Style importantStyle = sheet.Cells["B2"].GetStyle();
            importantStyle.Font.IsItalic = true;
            importantStyle.Font.Color = System.Drawing.Color.Red;
            importantStyle.Font.Underline = FontUnderlineType.Single;
            sheet.Cells["B2"].SetStyle(importantStyle);

            sheet.Cells["C3"].PutValue("Normal text");
            // No special style for C3 – will be rendered using default styling

            // Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save with inline styles only (DisableCss = true)
            htmlOptions.DisableCss = true;
            string inlinePath = "HtmlWithInlineStyles.html";
            workbook.Save(inlinePath, htmlOptions);

            // Save with external CSS (DisableCss = false)
            htmlOptions.DisableCss = false;
            string cssPath = "HtmlWithCssStyles.html";
            workbook.Save(cssPath, htmlOptions);

            // Retrieve file sizes
            long inlineSize = new FileInfo(inlinePath).Length;
            long cssSize = new FileInfo(cssPath).Length;

            // Output the comparison results
            Console.WriteLine($"File size with DisableCss = true (inline styles): {inlineSize} bytes");
            Console.WriteLine($"File size with DisableCss = false (CSS): {cssSize} bytes");

            // Indicate which approach yields a smaller file
            if (inlineSize < cssSize)
                Console.WriteLine("Inline styles produce a smaller HTML file.");
            else if (cssSize < inlineSize)
                Console.WriteLine("CSS styles produce a smaller HTML file.");
            else
                Console.WriteLine("Both files have the same size.");
        }
    }
}
