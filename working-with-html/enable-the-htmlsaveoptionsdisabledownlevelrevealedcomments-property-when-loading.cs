using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Create HTML save options and enable the property to disable downlevel-revealed comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableDownlevelRevealedComments = true;

        // Save the workbook as HTML using the configured options
        string outputFile = "output.html";
        workbook.Save(outputFile, htmlOptions);
    }
}