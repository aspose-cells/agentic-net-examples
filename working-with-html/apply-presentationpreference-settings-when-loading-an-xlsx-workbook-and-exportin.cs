using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to enable presentation‑preference rendering
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.PresentationPreference = true;   // more beautiful layout
        htmlOptions.IsFullPathLink = false;          // use relative links

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}