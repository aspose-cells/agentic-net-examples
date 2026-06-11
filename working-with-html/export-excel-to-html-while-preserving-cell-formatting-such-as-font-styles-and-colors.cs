using System;
using Aspose.Cells;

class ExportExcelToHtml
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Path where the HTML file will be saved
        string htmlFile = "output.html";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourceFile);

        // Configure HTML save options to keep all formatting
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export all data (including formulas, formatting, etc.)
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Preserve formulas in the HTML output
        htmlOptions.ExportFormula = true;

        // Export grid lines so the visual layout matches Excel
        htmlOptions.ExportGridLines = true;

        // Do not exclude unused styles – keeps every style definition,
        // ensuring that font styles and colors are retained exactly.
        htmlOptions.ExcludeUnusedStyles = false;

        // Save the workbook as an HTML file using the configured options
        workbook.Save(htmlFile, htmlOptions);
    }
}