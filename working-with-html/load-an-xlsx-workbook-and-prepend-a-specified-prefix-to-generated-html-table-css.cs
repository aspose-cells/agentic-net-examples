using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Input Excel file path
        string inputPath = "input.xlsx";

        // Output HTML file path
        string outputPath = "output.html";

        // Prefix to prepend to generated HTML table CSS IDs
        string cssIdPrefix = "myPrefix-";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and set the TableCssId prefix
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = cssIdPrefix;

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"HTML file saved with TableCssId prefix '{cssIdPrefix}'.");
    }
}