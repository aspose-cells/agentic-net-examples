using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook.
        // Replace "input.xlsx" with the path to your workbook file.
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and disable downlevel-revealed conditional comments.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableDownlevelRevealedComments = true;

        // Save the workbook as an HTML file using the configured options.
        // The output file will not contain downlevel-revealed comments.
        workbook.Save("output.html", htmlOptions);

        Console.WriteLine("Workbook has been saved to HTML with downlevel-revealed comments disabled.");
    }
}