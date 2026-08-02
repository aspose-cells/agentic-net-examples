using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");

        // Configure HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export row and column headings (e.g., A, B, 1, 2)
        htmlOptions.ExportRowColumnHeadings = true;

        // Define a CSS class that will be applied to heading cells.
        // Aspose.Cells generates the class name "heading" for row/column headings.
        htmlOptions.CssStyles = ".heading { font-weight: bold; color: #2A7AE2; }";

        // Optional: set a prefix for generated cell CSS classes (helps avoid name clashes)
        htmlOptions.CellCssPrefix = "c";

        // Save the workbook as an HTML file with the custom heading style
        workbook.Save("output_with_custom_headings.html", htmlOptions);
    }
}

// Author: Example demonstrating how to apply a custom CSS class to heading elements during HTML export.