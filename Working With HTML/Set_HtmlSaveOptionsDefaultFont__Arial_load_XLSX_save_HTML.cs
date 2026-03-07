using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and specify the default font
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Arial";

        // Save the workbook as HTML using the configured options
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}