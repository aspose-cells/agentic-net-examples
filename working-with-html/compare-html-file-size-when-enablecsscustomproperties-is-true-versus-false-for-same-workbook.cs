// Title: Compare Aspose.Cells HTML export size with EnableCssCustomProperties true vs false in C#
// AI Prompts: Generate a C# console program that creates a workbook, fills it with sample data, saves it to HTML twice—once with HtmlSaveOptions.EnableCssCustomProperties = true and once with it set to false—and prints the byte size of each file. | Write a .NET script that measures the impact of the EnableCssCustomProperties flag on the generated HTML by outputting a size comparison and a brief interpretation of the result.
// Common Searches: Aspose.Cells HTML export size difference when EnableCssCustomProperties is enabled | C# how to measure HTML file size with and without CSS custom properties using Aspose.Cells | Does setting EnableCssCustomProperties to true increase the size of generated HTML in Aspose.Cells? | Compare HTML output sizes for Aspose.Cells SaveOptions with CSS custom properties toggled | Performance impact of HtmlSaveOptions.EnableCssCustomProperties on HTML file size
// Tags: Aspose.Cells HtmlSaveOptions EnableCssCustomProperties comparison | C# measure HTML output size Aspose.Cells | HTML export size impact CSS custom properties Aspose.Cells | Aspose.Cells workbook to HTML file size analysis | EnableCssCustomProperties file size effect .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook with sample data, saves it to two HTML files—one with EnableCssCustomProperties set to true and another with it set to false—retrieves each file's byte length, and prints a comparison indicating whether the CSS custom properties setting increases, decreases, or leaves unchanged the HTML file size.
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

        // Define output file names
        string htmlWithCustomProps = "Workbook_WithCustomProperties.html";
        string htmlWithoutCustomProps = "Workbook_WithoutCustomProperties.html";

        // Save HTML with EnableCssCustomProperties = true
        HtmlSaveOptions optionsWith = new HtmlSaveOptions();
        optionsWith.EnableCssCustomProperties = true;
        workbook.Save(htmlWithCustomProps, optionsWith);

        // Save HTML with EnableCssCustomProperties = false
        HtmlSaveOptions optionsWithout = new HtmlSaveOptions();
        optionsWithout.EnableCssCustomProperties = false;
        workbook.Save(htmlWithoutCustomProps, optionsWithout);

        // Get file sizes
        long sizeWith = new FileInfo(htmlWithCustomProps).Length;
        long sizeWithout = new FileInfo(htmlWithoutCustomProps).Length;

        // Output the comparison results
        Console.WriteLine($"HTML size with EnableCssCustomProperties = true : {sizeWith} bytes");
        Console.WriteLine($"HTML size with EnableCssCustomProperties = false: {sizeWithout} bytes");

        // Simple comparison
        if (sizeWith > sizeWithout)
        {
            Console.WriteLine("Enabling CSS custom properties increases the HTML file size.");
        }
        else if (sizeWith < sizeWithout)
        {
            Console.WriteLine("Enabling CSS custom properties decreases the HTML file size.");
        }
        else
        {
            Console.WriteLine("Both settings produce HTML files of the same size.");
        }
    }
}
