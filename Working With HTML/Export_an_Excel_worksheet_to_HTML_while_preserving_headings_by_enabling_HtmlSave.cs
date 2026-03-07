using System;
using Aspose.Cells;

namespace ExportExcelToHtmlWithHeadings
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Path for the resulting HTML file
            string outputPath = "output_with_headings.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options and enable exporting of row/column headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportHeadings = true; // Obsolete property, kept for compatibility as per task

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML with headings at: {outputPath}");
        }
    }
}