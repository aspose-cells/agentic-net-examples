using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and set a custom TableCssId prefix
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table";

        // Path for the resulting HTML file
        string outputPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"HTML file saved to '{outputPath}' with TableCssId = '{saveOptions.TableCssId}'.");
    }
}