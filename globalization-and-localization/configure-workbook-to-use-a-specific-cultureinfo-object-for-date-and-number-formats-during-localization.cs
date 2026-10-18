// Title: Configure Aspose.Cells workbook to use French (fr-FR) CultureInfo for localized date and number formatting in C#
// Common Searches: C# how to apply French culture to Aspose.Cells workbook for date formatting | set workbook Settings.CultureInfo to fr-FR in Aspose.Cells example | localize numeric and date cells in Excel using Aspose.Cells and CultureInfo | Aspose.Cells date style respects workbook culture in C# | export Excel with French number format using Aspose.Cells C#
// Tags: Aspose.Cells workbook cultureinfo setting | French fr-FR localization Aspose.Cells | date formatting with CultureInfo in Excel C# | numeric formatting culture-specific Aspose.Cells | C# Excel export culture settings Aspose

using Aspose.Cells;
using System;
using System.Globalization;

// The sample creates a Workbook, assigns Settings.CultureInfo to French (fr-FR), adds a date and a numeric value, applies a short date style and a two‑decimal number style that follow the specified culture, and saves the file as LocalizedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure the workbook to use a specific CultureInfo (e.g., French - France)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Add sample data to demonstrate localized formatting
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a date value
        cells["A1"].PutValue(DateTime.Now);
        // Insert a numeric value
        cells["B1"].PutValue(12345.6789);

        // Apply a date style that respects the workbook's culture
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // Short date format
        cells["A1"].SetStyle(dateStyle);

        // Apply a numeric style that respects the workbook's culture
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Number = 2; // Number with two decimal places
        cells["B1"].SetStyle(numberStyle);

        // Save the workbook
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}
