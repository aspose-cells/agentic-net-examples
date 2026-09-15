// Title: Configure Aspose.Cells workbook to use French (fr-FR) culture for locale‑specific formatting of optional property values during export
// AI Prompts: Assign a specific CultureInfo (e.g., fr-FR) to Workbook.Settings.CultureInfo so that all date and numeric values are formatted according to that locale when the file is saved. | Create built‑in date and number styles that automatically follow the workbook's culture settings and export the workbook to an XLSX file.
// Common Searches: how to set cultureinfo for Aspose.Cells workbook export | Aspose.Cells format dates and numbers using French locale | export Excel with locale specific formatting Aspose.Cells .NET | configure workbook settings culture for optional property values Aspose.Cells
// Tags: Aspose.Cells workbook locale configuration | locale-aware number formatting Aspose.Cells | export Excel with French locale .NET | apply built-in date style using workbook culture | culture-specific optional property export Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// // This program creates a new workbook, sets its Settings.CultureInfo to French (fr-FR), inserts a date and a number, applies built‑in date and numeric styles that honor the culture, and saves the workbook as ExportedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Configure the workbook to use a specific culture (e.g., French - France)
        // This culture will be applied when formatting optional property values during export.
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Example data to demonstrate culture-specific formatting
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a date and a number that will be formatted according to the specified culture
        cells["A1"].PutValue(DateTime.Now);
        cells["A2"].PutValue(12345.67);

        // Apply a number format that relies on the workbook's culture settings
        Style style = workbook.CreateStyle();
        style.Number = 14; // Built‑in date format
        cells["A1"].SetStyle(style);

        style = workbook.CreateStyle();
        style.Number = 2; // Built‑in number format with two decimal places
        cells["A2"].SetStyle(style);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ExportedWorkbook.xlsx");
    }
}
