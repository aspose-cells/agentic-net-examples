// Title: Set a specific CultureInfo for date and number formatting in an Aspose.Cells workbook (C#)
// Common Searches: C# Aspose.Cells set workbook culture to French for date formatting | How to apply custom CultureInfo to Excel file using Aspose.Cells .NET | Localize numbers and dates in Aspose.Cells workbook with CultureInfo | Aspose.Cells workbook Settings.CultureInfo example code | Configure Excel globalization in Aspose.Cells C# project
// Tags: Aspose.Cells workbook cultureinfo configuration | Workbook.Settings.CultureInfo date localization | Excel number formatting French locale Aspose.Cells | C# Aspose.Cells globalization and localization | Apply CultureInfo to Excel workbook using Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a new Workbook, assigns a French (fr-FR) CultureInfo to the workbook's Settings.CultureInfo, writes a date and a numeric value so they are formatted according to that culture, and saves the file as LocalizedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define the specific culture (e.g., French - France)
        CultureInfo culture = new CultureInfo("fr-FR");

        // Apply the culture to the workbook settings for date and number formatting
        workbook.Settings.CultureInfo = culture;

        // Example data to demonstrate localization
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(DateTime.Now);   // Date will be formatted according to the culture
        sheet.Cells["B1"].PutValue(12345.67);       // Number will be formatted according to the culture

        // Save the workbook
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}
