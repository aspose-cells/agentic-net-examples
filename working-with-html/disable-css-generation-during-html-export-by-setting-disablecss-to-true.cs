using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Configure HTML save options to disable CSS generation (use inline styles only)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                DisableCss = true // Inline styles will be applied, no external CSS file will be created
            };

            // Save the workbook as HTML using the configured options (lifecycle: save)
            workbook.Save("OutputWithoutCss.html", htmlOptions);
        }
    }
}