// Title: Measure HTML file size reduction when disabling CSS with Aspose.Cells HtmlSaveOptions in C#
// Common Searches: C# Aspose.Cells HtmlSaveOptions DisableCss how much does it shrink HTML output | compare HTML output size with and without embedded CSS using Aspose.Cells | calculate markup size difference when saving a workbook to HTML with DisableCss true | Aspose.Cells generate HTML without CSS and get file size in bytes | measure HTML file size reduction by turning off CSS in Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions DisableCss | HTML size comparison Aspose.Cells | C# workbook to HTML without CSS | markup reduction measurement Aspose.Cells | memory stream HTML size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, and saves it twice to HTML using Aspose.Cells HtmlSaveOptions—once with CSS included (DisableCss = false) and once without CSS (DisableCss = true). It writes each HTML stream to a MemoryStream, reads the byte lengths, and prints both sizes along with the absolute and percentage reduction, demonstrating how disabling CSS reduces the generated markup.
class HtmlCssComparison
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with sample data to generate meaningful HTML markup
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Prepare HTML save options with DisableCss = false (default)
        HtmlSaveOptions optionsWithCss = new HtmlSaveOptions(SaveFormat.Html);
        optionsWithCss.DisableCss = false; // CSS will be included

        // Save to a memory stream to avoid file I/O
        MemoryStream htmlWithCssStream = new MemoryStream();
        workbook.Save(htmlWithCssStream, optionsWithCss);
        long sizeWithCss = htmlWithCssStream.Length;

        // Prepare HTML save options with DisableCss = true (no CSS)
        HtmlSaveOptions optionsWithoutCss = new HtmlSaveOptions(SaveFormat.Html);
        optionsWithoutCss.DisableCss = true; // CSS will be omitted

        // Save to another memory stream
        MemoryStream htmlWithoutCssStream = new MemoryStream();
        workbook.Save(htmlWithoutCssStream, optionsWithoutCss);
        long sizeWithoutCss = htmlWithoutCssStream.Length;

        // Output the file sizes for comparison
        Console.WriteLine($"HTML size with CSS:    {sizeWithCss} bytes");
        Console.WriteLine($"HTML size without CSS: {sizeWithoutCss} bytes");

        // Simple assessment of markup reduction
        long reduction = sizeWithCss - sizeWithoutCss;
        double reductionPercent = (double)reduction / sizeWithCss * 100;
        Console.WriteLine($"Markup reduction: {reduction} bytes ({reductionPercent:F2}%)");
    }
}
