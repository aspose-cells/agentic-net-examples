// Title: Export an Excel workbook to HTML with locale‑specific number formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, assigns Workbook.Settings.CultureInfo to a target locale (e.g., fr-FR), and saves the workbook as HTML with Aspose.Cells. | Demonstrate how to apply a built‑in numeric style that automatically follows the workbook’s culture before performing an HTML export in Aspose.Cells. | Explain the effect of changing the workbook’s CultureInfo on decimal and thousand separators in the HTML output produced by Aspose.Cells.
// Common Searches: how to export Excel to HTML with French number format using Aspose.Cells C# | Aspose.Cells set workbook culture before HTML conversion | C# locale specific decimal separator in HTML output from Excel | save workbook as HTML respecting regional settings Aspose.Cells | apply number format based on culture when converting Excel to HTML
// Tags: Aspose.Cells HTML export with culture settings | Workbook.Settings.CultureInfo number formatting | C# locale-aware Excel to HTML conversion | apply built-in numeric style Aspose.Cells | regional decimal separator HTML output

using System;
using System.Globalization;
using Aspose.Cells;

// Loads an Excel file, sets Workbook.Settings.CultureInfo to a specific locale (e.g., fr-FR) so numeric formats use locale‑appropriate separators, and saves the workbook as an HTML file using Aspose.Cells.
class ExportExcelToHtmlWithLocale
{
    static void Main()
    {
        // Load the existing Excel workbook
        // Replace "input.xlsx" with the path to your source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Set the workbook's culture to the desired locale (e.g., French - France)
        // This influences number formatting (decimal separators, thousand separators, etc.)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Optionally, ensure that cells use the built‑in number formats that respect the culture.
        // For example, apply a number format to a column if needed:
        // Worksheet sheet = workbook.Worksheets[0];
        // Style style = workbook.CreateStyle();
        // style.Number = 2; // Number format with two decimal places
        // StyleFlag flag = new StyleFlag { NumberFormat = true };
        // sheet.Cells["A1:A10"].ApplyStyle(style, flag);

        // Export the workbook to HTML
        // Replace "output.html" with the desired output path
        workbook.Save("output.html", SaveFormat.Html);
    }
}
