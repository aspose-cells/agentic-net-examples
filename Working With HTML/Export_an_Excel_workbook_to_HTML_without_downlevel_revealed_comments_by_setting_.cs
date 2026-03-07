using System;
using Aspose.Cells;

class ExportToHtml
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and disable downlevel-revealed comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableDownlevelRevealedComments = true;

        // Save the workbook as HTML using the configured options
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);
    }
}