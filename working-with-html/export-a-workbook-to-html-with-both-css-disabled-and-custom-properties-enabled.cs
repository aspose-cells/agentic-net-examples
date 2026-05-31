using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B2"].PutValue(12345);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Disable external CSS, use only inline styles
                DisableCss = true,

                // Enable CSS custom properties for optimized output
                EnableCssCustomProperties = true
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML at: {outputPath}");
        }
    }
}