using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to include cell comments in the output
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            IsExportComments = true   // Enable exporting of comments
        };

        // Save the workbook as an HTML file with comments included
        workbook.Save("output.html", htmlOptions);
    }
}