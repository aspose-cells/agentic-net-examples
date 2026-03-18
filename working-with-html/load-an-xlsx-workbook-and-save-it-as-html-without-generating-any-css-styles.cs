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

        // Create HTML save options and disable CSS generation
        // This forces Aspose.Cells to use only inline styles in the HTML output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true;

        // Path for the resulting HTML file
        string outputPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);
    }
}