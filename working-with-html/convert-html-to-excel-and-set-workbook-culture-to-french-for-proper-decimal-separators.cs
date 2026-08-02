// Title: Convert HTML to Excel with French Locale (fr-FR) using Aspose.Cells for .NET
// Description: This example shows how to load an HTML file into an Aspose.Cells Workbook with LoadOptions set to the French (fr-FR) CultureInfo, apply the same locale to the workbook settings, and save the result as an XLSX file so that commas are used for decimal separators and French date formats are respected.
// Keywords: Aspose.Cells | HTML to Excel | French locale | fr-FR | CultureInfo | decimal separator comma | .NET | LoadOptions | Workbook Settings | XLSX export
// Common Searches: Aspose.Cells load HTML with French locale | Convert HTML to XLSX using fr-FR culture | Set comma as decimal separator in Aspose.Cells | LoadOptions CultureInfo example .NET | HTML to Excel conversion French number format | Apply French date format when importing HTML
// Developer Intent: Load an HTML document into a workbook and export it to Excel while enforcing French numeric and date formatting.
// Use Cases: Transform French‑language financial reports delivered as HTML into Excel files that preserve correct decimal values. | Import tables from French web pages into Excel for analysis or reporting in a localized application. | Generate multilingual Excel outputs from HTML templates, automatically applying French number and date conventions. | Batch‑process a directory of HTML files, converting each to XLSX with French locale settings for data migration projects. | Create Excel dashboards from HTML sources where French users expect commas as decimal separators.
// AI Prompts: Provide code to set a custom number format that forces a comma as the decimal separator after loading HTML with Aspose.Cells. | Show how to iterate over all HTML files in a folder, convert each to XLSX using fr-FR culture, and log any conversion errors. | Explain how to handle French date strings (e.g., "31/12/2023") when loading HTML into an Aspose.Cells workbook.

using System;
using System.Globalization;
using Aspose.Cells;

// This example shows how to load an HTML file into an Aspose.Cells Workbook with LoadOptions set to the French (fr-FR) CultureInfo, apply the same locale to the workbook settings, and save the result as an XLSX file so that commas are used for decimal separators and French date formats are respected.
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

        // Load the HTML file into a workbook using the specified load options
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Ensure the workbook's settings also use French culture (optional but reinforces the setting)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Save the workbook as an Excel file (XLSX)
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
