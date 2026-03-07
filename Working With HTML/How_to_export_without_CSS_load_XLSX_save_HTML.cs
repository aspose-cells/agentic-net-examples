using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and disable CSS generation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true; // Use only inline styles

        // Save the workbook as HTML without external CSS files
        workbook.Save("output.html", htmlOptions);
    }
}