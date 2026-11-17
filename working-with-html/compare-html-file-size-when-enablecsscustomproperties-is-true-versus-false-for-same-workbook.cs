// Title: Compare HTML export size with EnableCssCustomProperties set to true versus false using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# compare HTML file size when EnableCssCustomProperties is true | measure impact of EnableCssCustomProperties on HTML export size with Aspose.Cells | C# Aspose.Cells HtmlSaveOptions EnableCssCustomProperties false size difference | determine HTML byte size difference with CSS custom properties enabled in Aspose.Cells
// Tags: HtmlSaveOptions EnableCssCustomProperties performance | Aspose.Cells HTML export size analysis | C# workbook to HTML byte size measurement | CSS custom properties impact on Aspose.Cells HTML output | memory stream size comparison Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a 100‑row by 10‑column worksheet, saves it to HTML twice—once with HtmlSaveOptions.EnableCssCustomProperties set to true and once set to false—captures the byte length of each MemoryStream, and prints both sizes, allowing developers to see how CSS custom properties affect the generated HTML file size.
class HtmlCssCustomPropertiesComparison
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate the worksheet with sample data
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Prepare HTML save options with EnableCssCustomProperties = true
        HtmlSaveOptions optionsTrue = new HtmlSaveOptions(SaveFormat.Html);
        optionsTrue.EnableCssCustomProperties = true;

        // Prepare HTML save options with EnableCssCustomProperties = false
        HtmlSaveOptions optionsFalse = new HtmlSaveOptions(SaveFormat.Html);
        optionsFalse.EnableCssCustomProperties = false;

        // Save the workbook to memory streams using both option sets
        long sizeWithCustomProperties;
        long sizeWithoutCustomProperties;

        using (MemoryStream msTrue = new MemoryStream())
        {
            workbook.Save(msTrue, optionsTrue);
            sizeWithCustomProperties = msTrue.Length;
        }

        using (MemoryStream msFalse = new MemoryStream())
        {
            workbook.Save(msFalse, optionsFalse);
            sizeWithoutCustomProperties = msFalse.Length;
        }

        // Output the file sizes for comparison
        Console.WriteLine($"HTML size with EnableCssCustomProperties = true : {sizeWithCustomProperties} bytes");
        Console.WriteLine($"HTML size with EnableCssCustomProperties = false: {sizeWithoutCustomProperties} bytes");
    }
}
