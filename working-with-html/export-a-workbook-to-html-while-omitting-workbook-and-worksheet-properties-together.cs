using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Configure HTML save options to omit workbook and worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportWorkbookProperties = false,   // Do not export workbook properties
                ExportWorksheetProperties = false   // Do not export worksheet properties
            };

            // Save the workbook as HTML using the configured options
            string outputPath = "output_without_properties.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML without properties at: {outputPath}");
        }
    }
}