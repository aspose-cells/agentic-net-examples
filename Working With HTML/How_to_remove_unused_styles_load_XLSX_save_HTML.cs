using System;
using Aspose.Cells;

class RemoveUnusedStylesExample
{
    static void Main()
    {
        // Load the XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Remove all styles that are not used in the workbook
        workbook.RemoveUnusedStyles();

        // Configure HTML save options to exclude unused styles (default is true)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExcludeUnusedStyles = true;

        // Save the workbook as HTML
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("HTML file saved with unused styles removed.");
    }
}