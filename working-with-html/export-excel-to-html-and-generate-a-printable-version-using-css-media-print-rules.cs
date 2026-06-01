using System;
using Aspose.Cells;

class ExportExcelToHtmlPrintable
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define custom CSS that includes @media print rules.
        // The rules hide elements with class "no-print" and force a page break after each table when printing.
        string customCss = @"
            @media print {
                body { margin:0; }
                table { page-break-after: always; }
                .no-print { display:none; }
            }";

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Save the HTML as a single file so that CssStyles is applied.
        htmlOptions.SaveAsSingleFile = true;
        // Inject the custom CSS (including the @media print block)
        htmlOptions.CssStyles = customCss;

        // Save the workbook as an HTML file. The generated HTML contains the print‑specific CSS.
        workbook.Save("output.html", htmlOptions);
    }
}