// Title: Set a specific CultureInfo on an Aspose.Cells workbook to localize date and number formats in C#
// AI Prompts: Assign workbook.Settings.CultureInfo to a target locale before inserting data so that dates and numbers adopt that culture's formatting. | Create a new workbook, apply French (fr-FR) culture, write a DateTime value and a decimal, then save the file as an .xlsx document. | Change the workbook's CultureInfo to German (de-DE) after creation and retrieve the formatted cell strings to confirm locale‑specific rendering.
// Common Searches: Aspose.Cells set workbook cultureinfo for French localization C# | How to apply a specific CultureInfo to an Excel workbook using Aspose.Cells .NET | Localize date and number formatting in an Aspose.Cells workbook | C# Aspose.Cells change workbook locale to German (de-DE) | Saving an Aspose.Cells workbook with culture‑specific formats
// Tags: set workbook cultureinfo Aspose.Cells | localize date formatting Aspose.Cells C# | apply cultureinfo to Excel workbook | culture-specific number formatting Aspose.Cells | Aspose.Cells workbook localization fr-FR

using Aspose.Cells;
using System;
using System.Globalization;

// The example creates a Workbook, sets its Settings.CultureInfo to a chosen locale (e.g., fr-FR), adds a date and a numeric value to demonstrate localized formatting, and saves the workbook as LocalizedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Configure the workbook to use a specific CultureInfo (e.g., French - France)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Add sample data to demonstrate localized formatting
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(DateTime.Now);   // Date will be formatted according to the culture
        sheet.Cells["A2"].PutValue(12345.67);       // Number will be formatted according to the culture

        // Save the workbook
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}
