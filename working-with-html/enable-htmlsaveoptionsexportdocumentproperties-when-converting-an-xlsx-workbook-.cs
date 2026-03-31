using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportDocumentPropertiesDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set built‑in document properties that will be exported to HTML
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Document";

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World!");

            // Create HTML save options and explicitly enable ExportDocumentProperties
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportDocumentProperties = true; // include document metadata in the HTML output

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", options);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportDocumentPropertiesDemo.Run();
        }
    }
}