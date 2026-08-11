// Title: Export Excel to HTML with French locale number formatting using Aspose.Cells for .NET
// Description: Demonstrates how to set a workbook's CultureInfo to French (fr-FR) in Aspose.Cells, apply a built‑in numeric format, configure HtmlSaveOptions, and save the sheet as HTML so numbers appear with French thousand and decimal separators.
// Keywords: Aspose.Cells | HTML export | locale number formatting | French culture | C# | Workbook CultureInfo | HtmlSaveOptions | Excel to HTML | number format #,##0.00 | .NET
// Common Searches: Aspose.Cells export Excel to HTML with French formatting | set workbook culture for HTML output Aspose.Cells | locale specific number format in HTML export .NET | how to apply French number format when saving as HTML | HtmlSaveOptions culture info Aspose.Cells
// Developer Intent: Generate an HTML file from an Excel workbook where numeric cells follow the French (fr-FR) number formatting rules.
// Use Cases: Create web‑ready financial reports for French‑speaking audiences. | Automate multi‑regional dashboards that display numbers with correct local separators. | Produce HTML versions of spreadsheets while preserving culture‑aware numeric formatting.
// AI Prompts: Show me C# code to export an Aspose.Cells workbook to HTML using German (de-DE) number formatting. | How can I keep custom number formats when saving an Excel sheet as HTML with Aspose.Cells? | Explain the steps to change a workbook's culture to Japanese and export it to HTML in .NET.

using System.Globalization;
using Aspose.Cells;

// Demonstrates how to set a workbook's CultureInfo to French (fr-FR) in Aspose.Cells, apply a built‑in numeric format, configure HtmlSaveOptions, and save the sheet as HTML so numbers appear with French thousand and decimal separators.
class ExportExcelToHtmlWithLocale
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook culture to French (France) – this will affect number formatting
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Get the first worksheet and put a numeric value
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1234.56);

        // Apply a number format that uses thousand separator and two decimal places
        // Built‑in format 10 corresponds to "#,##0.00"
        Style style = sheet.Cells["A1"].GetStyle();
        style.Number = 10;
        sheet.Cells["A1"].SetStyle(style);

        // Configure HTML save options (export all data)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the workbook as HTML; the numbers will be formatted according to the French locale
        workbook.Save("ExportedWithLocale.html", htmlOptions);
    }
}
