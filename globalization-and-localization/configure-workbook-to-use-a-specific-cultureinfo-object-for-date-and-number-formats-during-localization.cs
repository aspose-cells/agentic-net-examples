// Title: Set Workbook CultureInfo (de-DE) for Locale‑Aware Number and Date Formatting in Aspose.Cells C#
// Description: Shows how to assign a specific CultureInfo (German – de-DE) to Aspose.Cells Workbook.Settings.CultureInfo, then apply custom numeric and date patterns that automatically follow the locale’s separators, month names, and date order, and finally save the localized Excel file.
// Keywords: Aspose.Cells | C# | CultureInfo | de-DE | German locale | number formatting | date formatting | Excel localization | workbook.Settings.CultureInfo | locale-aware formatting
// Common Searches: Aspose.Cells set workbook cultureinfo | C# Aspose.Cells German number format | How to localize dates in Aspose.Cells | Workbook.Settings.CultureInfo example | Apply de-DE culture to Excel with Aspose | Locale specific formatting Aspose.Cells
// Developer Intent: Configure a workbook to use a chosen CultureInfo so that numeric and date cells are automatically formatted according to that locale.
// Use Cases: Generate a German‑language financial report where currency values and dates follow de‑DE conventions without manual string handling. | Create an invoice template that adapts number separators and month names based on the workbook’s CultureInfo for multiple European markets. | Export data to Excel for users in Germany, ensuring dates appear as dd MMMM yyyy and numbers use commas as decimal separators. | Produce a multi‑regional dashboard where each worksheet can be assigned a different CultureInfo for region‑specific formatting.
// AI Prompts: Provide C# code to change the workbook culture to French (fr-FR) and update a custom number format using Aspose.Cells. | Write a script that reads the current CultureInfo of a workbook, then prints the formatted values of a sample number and date for both de-DE and en-US locales. | Explain how to assign distinct CultureInfo objects to individual worksheets within the same workbook for mixed‑region reporting in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;

// Shows how to assign a specific CultureInfo (German – de-DE) to Aspose.Cells Workbook.Settings.CultureInfo, then apply custom numeric and date patterns that automatically follow the locale’s separators, month names, and date order, and finally save the localized Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the workbook's culture to German (Germany).
        // This culture will be used for number and date formatting.
        workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------
        // Demonstrate number formatting
        // -------------------------
        Cell numberCell = sheet.Cells["A1"];
        numberCell.PutValue(1234567.89); // Put a numeric value

        // Apply a custom number format; the actual separators will follow the German culture
        Style numberStyle = numberCell.GetStyle();
        numberStyle.Custom = "#,##0.00";
        numberCell.SetStyle(numberStyle);

        // -------------------------
        // Demonstrate date formatting
        // -------------------------
        Cell dateCell = sheet.Cells["A2"];
        dateCell.PutValue(DateTime.Now); // Put a date/time value

        // Apply a custom date format; month names and order will follow the German culture
        Style dateStyle = dateCell.GetStyle();
        dateStyle.Custom = "dd MMMM yyyy";
        dateCell.SetStyle(dateStyle);

        // Output the culture‑adjusted formatted strings to the console
        Console.WriteLine("Formatted number (de-DE): " + numberCell.StringValue);
        Console.WriteLine("Formatted date   (de-DE): " + dateCell.StringValue);

        // Save the workbook to a file
        workbook.Save("LocalizedWorkbook.xlsx");
    }
}
