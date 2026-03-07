using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable gridline export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}