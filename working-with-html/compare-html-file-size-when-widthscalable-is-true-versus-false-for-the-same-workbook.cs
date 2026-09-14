// Title: Compare HTML output file sizes with HtmlSaveOptions.WidthScalable set to true and false in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, saves it to HTML twice—once with HtmlSaveOptions.WidthScalable = true and once with false—and prints the byte size of each generated file. | Extend the sample to compute the size difference between the two HTML files and display which WidthScalable setting produces the larger file.
// Common Searches: Aspose.Cells how does WidthScalable affect HTML file size in C# | C# compare size of HTML files generated with WidthScalable true versus false | measure Aspose.Cells HTML export size when using scalable width option | difference in HTML output size with HtmlSaveOptions.WidthScalable true in .NET
// Tags: Aspose.Cells HtmlSaveOptions WidthScalable | C# export workbook to HTML size comparison | HTML file size impact of WidthScalable setting | measure Aspose.Cells HTML output size

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, saves it to two HTML files using HtmlSaveOptions with WidthScalable set to true and false, reads each file's byte length, prints the sizes, and reports which setting generates a larger HTML file.
class HtmlWidthScalableComparison
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Define file names for the two HTML outputs
        string htmlTruePath = "Workbook_WidthScalable_True.html";
        string htmlFalsePath = "Workbook_WidthScalable_False.html";

        // Save with WidthScalable = true
        HtmlSaveOptions optionsTrue = new HtmlSaveOptions();
        optionsTrue.WidthScalable = true;
        workbook.Save(htmlTruePath, optionsTrue);

        // Save with WidthScalable = false
        HtmlSaveOptions optionsFalse = new HtmlSaveOptions();
        optionsFalse.WidthScalable = false;
        workbook.Save(htmlFalsePath, optionsFalse);

        // Get file sizes
        long sizeTrue = new FileInfo(htmlTruePath).Length;
        long sizeFalse = new FileInfo(htmlFalsePath).Length;

        // Output the comparison results
        Console.WriteLine($"HTML size with WidthScalable = true : {sizeTrue} bytes");
        Console.WriteLine($"HTML size with WidthScalable = false: {sizeFalse} bytes");

        if (sizeTrue > sizeFalse)
        {
            Console.WriteLine("WidthScalable = true produces a larger HTML file.");
        }
        else if (sizeTrue < sizeFalse)
        {
            Console.WriteLine("WidthScalable = false produces a larger HTML file.");
        }
        else
        {
            Console.WriteLine("Both settings produce HTML files of the same size.");
        }
    }
}
