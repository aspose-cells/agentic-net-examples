using System;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Contains HtmlSaveOptions

class ExportConditionalFormattingToHtml
{
    static void Main()
    {
        // Path to the source Excel file that contains DataBar, ColorScale, and IconSet CF.
        string sourceFile = "input.xlsx";

        // Path for the generated HTML file.
        string htmlFile = "output.html";

        // Load the workbook from the Excel file.
        Workbook workbook = new Workbook(sourceFile);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Export all data (including conditional formatting) to HTML.
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the workbook as HTML with the specified options.
        workbook.Save(htmlFile, htmlOptions);

        Console.WriteLine("HTML file with DataBar, ColorScale, and IconSet exported successfully.");
    }
}