using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SampleSheet";

        // Populate some data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue("Data1");
        sheet.Cells["B2"].PutValue("Data2");

        // Define a page header that contains an <h1> tag.
        // When ExportPageHeaders is enabled, this header will appear as an <h1> element in the HTML output.
        sheet.PageSetup.SetHeader(0, "<h1>Sample Worksheet</h1>");

        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            // Export row (1,2,…) and column (A,B,…) headings.
            ExportRowColumnHeadings = true,

            // Include the page header defined above.
            ExportPageHeaders = true,

            // Optional: set the HTML page title.
            PageTitle = "Worksheet with Headings"
        };

        // Save the workbook as an HTML file with the specified options.
        workbook.Save("WorksheetWithHeadings.html", options);
    }
}