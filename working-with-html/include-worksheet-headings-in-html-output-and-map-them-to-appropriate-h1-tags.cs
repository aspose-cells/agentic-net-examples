using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add column headings (these will be rendered as <h1>‑like elements in the HTML)
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["C1"].PutValue("Quantity");

        // Add some sample data rows
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1.20);
        worksheet.Cells["C2"].PutValue(50);

        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(0.80);
        worksheet.Cells["C3"].PutValue(30);

        // Configure HTML save options to export row/column headings.
        // When ExportRowColumnHeadings is true, the first row (our headings) will be output
        // as HTML table header cells, which browsers render similarly to <h1> tags.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportRowColumnHeadings = true
        };

        // Save the workbook as an HTML file with headings included
        workbook.Save("WorksheetWithHeadings.html", htmlOptions);
    }
}

// Author: Aspose.Cells .NET example – includes worksheet headings in HTML output.