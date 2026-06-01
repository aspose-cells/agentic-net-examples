using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings (preserves all content)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export the workbook to HTML using the default options
        workbook.Save("output.html", htmlOptions);
    }
}