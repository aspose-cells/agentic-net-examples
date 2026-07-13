using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingVerification
{
    // Author: Aspose.Cells .NET example – verifies worksheet title appears as a heading in exported HTML
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom worksheet name (this will be the title we expect in HTML)
            string worksheetTitle = "SalesReport";
            sheet.Name = worksheetTitle;

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(200);

            // Configure HTML save options to export row/column headings and worksheet properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRowColumnHeadings = true,   // export A, B, 1, 2 … headings
                ExportWorksheetProperties = true // include worksheet name as a heading
            };

            // Save the workbook as HTML
            string htmlPath = "SalesReport.html";
            workbook.Save(htmlPath, htmlOptions);

            // Read the generated HTML file
            string htmlContent = File.ReadAllText(htmlPath);

            // Simple verification: check if the worksheet title appears inside an <h1> tag
            // (Aspose.Cells typically wraps the worksheet name in a <h1> element)
            string expectedHeading = $"<h1>{worksheetTitle}</h1>";
            bool headingFound = htmlContent.Contains(expectedHeading, StringComparison.OrdinalIgnoreCase);

            // Output verification result
            Console.WriteLine(headingFound
                ? "Verification succeeded: worksheet title heading found in HTML."
                : "Verification failed: worksheet title heading not found in HTML.");
        }
    }
}