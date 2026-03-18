using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to suppress downlevel‑revealed comments
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableDownlevelRevealedComments = true;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}