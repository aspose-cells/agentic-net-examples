using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Optionally ensure the worksheet's gridlines are visible
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.IsGridlinesVisible = true;

        // Create HTML save options and enable gridline export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true
        };

        // Save the workbook as an HTML file with gridlines displayed
        workbook.Save("output.html", htmlOptions);
    }
}