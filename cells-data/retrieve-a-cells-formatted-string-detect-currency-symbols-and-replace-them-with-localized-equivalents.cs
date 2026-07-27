using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCurrencyLocalization
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the workbook culture to French (France) where the currency symbol is €
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare a cell with a numeric value and a built‑in currency format (Number = 164)
            Cell sourceCell = cells["A1"];
            sourceCell.PutValue(1234.56);
            Style style = sourceCell.GetStyle();
            style.Number = 164; // Currency format (e.g., $1,234.56 in en-US)
            sourceCell.SetStyle(style);

            // Retrieve the formatted string as shown in Excel (includes the default $ symbol)
            string formatted = sourceCell.DisplayStringValue; // e.g., "$1,234.56"

            // Detect the currency symbol used in the current culture
            string localCurrencySymbol = workbook.Settings.CultureInfo.NumberFormat.CurrencySymbol; // e.g., "€"

            // Replace the generic "$" (or any non‑local symbol) with the localized currency symbol
            // Here we simply replace the first character if it is a known currency placeholder.
            // For a more robust solution you could use regex to detect any non‑digit symbols.
            string localized = formatted.Replace("$", localCurrencySymbol);

            // Write the localized string back to another cell for demonstration
            Cell resultCell = cells["B1"];
            resultCell.PutValue(localized);

            // Save the workbook (lifecycle: save)
            workbook.Save("CurrencyLocalized.xlsx");
        }
    }
}