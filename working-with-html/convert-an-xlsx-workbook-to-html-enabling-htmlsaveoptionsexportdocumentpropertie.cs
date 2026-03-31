using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToHtmlWithProperties
    {
        public static void Run()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Enable exporting of document, workbook, and worksheet properties
            options.ExportDocumentProperties = true;   // Document properties
            options.ExportWorkbookProperties = true;   // Workbook properties
            options.ExportWorksheetProperties = true; // Worksheet properties

            // Save the workbook as an HTML file with the specified options
            workbook.Save("output.html", options);

            Console.WriteLine("Workbook successfully converted to HTML with metadata exported.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertXlsxToHtmlWithProperties.Run();
        }
    }
}