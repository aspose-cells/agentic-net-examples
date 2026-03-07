using System;
using Aspose.Cells;

namespace ExcelToHtmlWithHeadings
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel file (XLSX)
            Workbook workbook = new Workbook("source.xlsx");

            // Configure HTML save options to include row and column headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportRowColumnHeadings = true;   // Export A, B, ... and 1, 2, ...

            // Save the workbook as an HTML file with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Excel file has been successfully converted to HTML with headings.");
        }
    }
}