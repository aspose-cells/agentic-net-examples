// Title: Compare HTML output file size with WidthScalable true vs false using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# measure HTML file size difference with WidthScalable true | How to compare HTML output size when HtmlSaveOptions.WidthScalable is false in .NET | C# code to get size of HTML saved workbook using WidthScalable option | Effect of WidthScalable on HTML file size in Aspose.Cells for .NET | Determine which WidthScalable setting produces larger HTML in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions WidthScalable setting | HTML output size analysis Aspose.Cells | C# workbook to HTML scaling option | compare generated HTML file sizes Aspose.Cells | measure HTML size with WidthScalable true

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, saves it twice as HTML—once with HtmlSaveOptions.WidthScalable set to true and once to false—then reads the file sizes and reports which setting generates the larger HTML file.
class HtmlWidthScalableComparison
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data to make the HTML output non‑trivial
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define temporary file paths
        string pathTrue = Path.Combine(Path.GetTempPath(), "Workbook_WidthScalable_True.html");
        string pathFalse = Path.Combine(Path.GetTempPath(), "Workbook_WidthScalable_False.html");

        // Save HTML with WidthScalable = true
        HtmlSaveOptions optionsTrue = new HtmlSaveOptions(SaveFormat.Html);
        optionsTrue.WidthScalable = true;
        workbook.Save(pathTrue, optionsTrue);

        // Save HTML with WidthScalable = false
        HtmlSaveOptions optionsFalse = new HtmlSaveOptions(SaveFormat.Html);
        optionsFalse.WidthScalable = false;
        workbook.Save(pathFalse, optionsFalse);

        // Get file sizes
        long sizeTrue = new FileInfo(pathTrue).Length;
        long sizeFalse = new FileInfo(pathFalse).Length;

        // Output comparison results
        Console.WriteLine($"HTML size with WidthScalable = true : {sizeTrue} bytes");
        Console.WriteLine($"HTML size with WidthScalable = false: {sizeFalse} bytes");

        if (sizeTrue > sizeFalse)
            Console.WriteLine("WidthScalable = true produces a larger HTML file.");
        else if (sizeTrue < sizeFalse)
            Console.WriteLine("WidthScalable = false produces a larger HTML file.");
        else
            Console.WriteLine("Both settings produce HTML files of the same size.");

        // Clean up temporary files (optional)
        // File.Delete(pathTrue);
        // File.Delete(pathFalse);
    }
}
