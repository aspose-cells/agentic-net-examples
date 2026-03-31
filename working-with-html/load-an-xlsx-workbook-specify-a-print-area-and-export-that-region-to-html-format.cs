using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaToHtml
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any desired worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the print area that should be exported (e.g., B2 to D5)
            worksheet.PageSetup.PrintArea = "B2:D5";

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,   // Export only the print area
                ExportGridLines = true        // Optional: include grid lines in the HTML
            };

            // Save the specified region as an HTML file
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Print area exported to HTML successfully.");
        }
    }
}