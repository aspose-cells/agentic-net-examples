// Title: Export Excel to HTML with German (de‑DE) Number Formatting using Aspose.Cells for .NET
// Description: Demonstrates how to set a workbook's CultureInfo to de‑DE, apply a locale‑aware numeric style, configure HtmlSaveOptions with ExportDataOptions.All, and save the sheet as an HTML file that displays German‑style separators.
// Keywords: Aspose.Cells | HTML export | locale specific number format | workbook culture | de-DE | C# .NET | HtmlSaveOptions | ExportDataOptions.All | custom numeric format | Excel to HTML
// Common Searches: Aspose.Cells export Excel to HTML with German formatting | set workbook culture de-DE Aspose.Cells | locale aware number format in HTML export .NET | HtmlSaveOptions ExportDataOptions.All example | preserve Excel culture in HTML output
// Developer Intent: Create an HTML representation of an Excel workbook that automatically uses German numeric separators.
// Use Cases: Produce web‑ready reports for German audiences where currency and decimal values follow local conventions. | Render an Excel worksheet as HTML while keeping hidden rows/columns visible via ExportDataOptions.All. | Automate batch conversion of localized Excel files to HTML for multilingual portals.
// AI Prompts: Show how to export an Excel workbook to HTML with French (fr‑FR) culture using Aspose.Cells. | Provide C# code to apply a locale‑aware custom number format to a range before saving as HTML. | Explain the impact of HtmlSaveOptions.ExportDataOptions on the generated HTML in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to set a workbook's CultureInfo to de‑DE, apply a locale‑aware numeric style, configure HtmlSaveOptions with ExportDataOptions.All, and save the sheet as an HTML file that displays German‑style separators.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook culture to German (Germany) – this will affect number formatting
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into a cell
            sheet.Cells["A1"].PutValue(12345.67);

            // Apply a numeric format that respects the culture (e.g., thousand separator, decimal separator)
            Style style = sheet.Cells["A1"].GetStyle();
            style.Custom = "#,##0.00";   // Custom format; separators will follow the set culture
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options (export all data)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDataOptions = HtmlExportDataOptions.All
            };

            // Save the workbook as an HTML file
            workbook.Save("ExportedWithLocale.html", htmlOptions);
        }
    }
}
