using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options to use only inline styles (no external CSS)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true; // Inline styling only

        // Save the workbook as HTML
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}