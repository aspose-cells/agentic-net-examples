// Title: Convert HTML to Excel with French Locale (fr-FR) using Aspose.Cells for .NET
// Description: Loads an HTML document into an Aspose.Cells Workbook with LoadOptions.CultureInfo set to fr-FR, applies the same locale to workbook settings, and saves the result as an XLSX file so numbers use a comma as the decimal separator.
// Keywords: Aspose.Cells | HTML to Excel | French locale | fr-FR | CultureInfo | .NET | comma decimal separator | LoadOptions | SaveFormat.Xlsx
// Common Searches: Aspose.Cells load HTML with French culture | C# convert HTML table to Excel using fr-FR | set comma decimal separator when converting HTML to XLSX | LoadOptions CultureInfo example Aspose.Cells | preserve French number formatting in Excel export
// Developer Intent: Load an HTML file into a workbook with French regional settings and export it as an XLSX workbook.
// Use Cases: Transform financial HTML reports that use commas for decimals into Excel files for French‑speaking analysts. | Generate Excel worksheets from web pages for a French audience while maintaining correct numeric formatting. | Batch‑process multiple HTML tables into XLSX files with French locale to meet regulatory reporting requirements.
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file with CultureInfo set to fr-FR and save it as an XLSX workbook. | Explain how LoadOptions.CultureInfo influences number parsing during HTML‑to‑Excel conversion in Aspose.Cells. | Provide a step‑by‑step guide for batch converting HTML files to Excel while applying French decimal formatting.

using System;
using System.Globalization;
using Aspose.Cells;

// Loads an HTML document into an Aspose.Cells Workbook with LoadOptions.CultureInfo set to fr-FR, applies the same locale to workbook settings, and saves the result as an XLSX file so numbers use a comma as the decimal separator.
class HtmlToExcelFrenchCulture
{
    static void Main()
    {
        // Source HTML file and target Excel file paths
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        // Create load options for HTML format and set French culture (uses comma as decimal separator)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        loadOptions.CultureInfo = new CultureInfo("fr-FR");

        // Load the HTML file into a workbook using the specified culture
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Ensure workbook settings also use French culture (optional but reinforces the setting)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Save the workbook as an Excel file (XLSX)
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
