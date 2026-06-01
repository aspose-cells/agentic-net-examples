using System;
using System.Globalization;
using Aspose.Cells;

class CurrencyLocalizationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook's culture to German (Euro) for demonstration
        wb.Settings.CultureInfo = new CultureInfo("de-DE");

        // Access the first worksheet and a target cell
        Worksheet ws = wb.Worksheets[0];
        Cell cell = ws.Cells["A1"];

        // Put a numeric value into the cell
        cell.PutValue(1234.56);

        // Apply a built‑in currency number format to the cell
        Style style = cell.GetStyle();
        style.Number = 164; // Currency format
        cell.SetStyle(style);

        // Retrieve the formatted string as it would appear in Excel
        string formattedString = cell.DisplayStringValue;

        // Determine the currency symbol for the workbook's culture
        string localeCurrencySymbol = ((CultureInfo)wb.Settings.CultureInfo).NumberFormat.CurrencySymbol;

        // Replace common currency symbols with the locale‑specific one
        // Extend the replacement list as needed for other symbols
        string localizedString = formattedString
            .Replace("$", localeCurrencySymbol)
            .Replace("USD", localeCurrencySymbol)
            .Replace("€", localeCurrencySymbol)
            .Replace("£", localeCurrencySymbol);

        // Output the original and localized strings
        Console.WriteLine("Original formatted string: " + formattedString);
        Console.WriteLine("Localized string: " + localizedString);

        // Save the workbook (optional)
        wb.Save("CurrencyLocalizationDemo.xlsx");
    }
}