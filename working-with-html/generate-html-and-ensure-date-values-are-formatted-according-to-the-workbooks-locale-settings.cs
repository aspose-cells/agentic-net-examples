// Title: C# – Export Aspose.Cells Workbook to HTML with Locale‑Aware Date Formatting
// Description: Demonstrates how to create a workbook, assign a regional setting (e.g., United Kingdom), insert a DateTime value, apply the built‑in short date style, configure HtmlSaveOptions for full data export and HTML5 output, and save the file as HTML so that dates are rendered according to the workbook’s locale.
// Keywords: Aspose.Cells | C# | HTML export | locale date format | workbook region | HtmlSaveOptions | Html5 output | United Kingdom regional settings | short date style | regional formatting
// Common Searches: Aspose.Cells export HTML date format UK | C# set workbook region for HTML output | How to format dates by locale when saving as HTML with Aspose.Cells | HtmlSaveOptions regional settings example | Export spreadsheet to HTML with localized dates
// Developer Intent: Create an HTML representation of a workbook where every date cell respects the workbook’s regional settings.
// Use Cases: Publish a web‑based financial report that automatically shows dates in the target audience’s local format. | Generate localized HTML invoices where the date appears in the format defined by the workbook’s region. | Provide an online spreadsheet preview that preserves regional date conventions without additional client‑side scripting.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with dates formatted according to the workbook’s Region property. | Show how to change the workbook’s regional setting to French (France) and save it as HTML so dates appear in the French short date format. | Explain how HtmlSaveOptions interacts with workbook regional settings to produce locale‑specific date strings in the HTML output.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, assign a regional setting (e.g., United Kingdom), insert a DateTime value, apply the built‑in short date style, configure HtmlSaveOptions for full data export and HTML5 output, and save the file as HTML so that dates are rendered according to the workbook’s locale.
class GenerateHtmlWithLocaleDate
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's regional settings (example: United Kingdom)
        workbook.Settings.Region = CountryCode.UnitedKingdom;

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a date value into cell A1
        worksheet.Cells["A1"].PutValue(new DateTime(2023, 12, 31));

        // Apply a built‑in date format (number format 14) which respects the region setting
        Style dateStyle = worksheet.Cells["A1"].GetStyle();
        dateStyle.Number = 14; // Built‑in short date format
        worksheet.Cells["A1"].SetStyle(dateStyle);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export all data (including styles and formatting)
            ExportDataOptions = HtmlExportDataOptions.All,
            // Use HTML5 standard for the output
            HtmlVersion = HtmlVersion.Html5,
            // No need to calculate formulas for this example
            CalculateFormula = false
        };

        // Save the workbook as an HTML file; the date will be formatted according to the workbook's locale
        workbook.Save("output.html", saveOptions);
    }
}
