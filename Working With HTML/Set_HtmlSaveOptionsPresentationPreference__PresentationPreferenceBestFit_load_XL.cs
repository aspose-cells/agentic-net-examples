using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable presentation‑preference (best‑fit) for a more beautiful HTML output
        htmlOptions.PresentationPreference = true;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}