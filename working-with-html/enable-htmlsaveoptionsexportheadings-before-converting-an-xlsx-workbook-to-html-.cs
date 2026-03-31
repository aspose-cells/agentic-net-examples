using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Initialize HTML save options and enable exporting of row/column headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportRowColumnHeadings = true; // Retain worksheet headings in the HTML output

            // Save the workbook as an HTML file with headings enabled
            workbook.Save("WorkbookWithHeadings.html", htmlOptions);
        }
    }
}