// Title: Aspose.Cells for .NET – Retrieve a cell's DisplayStringValue and localize its currency symbol
// Description: Load an Excel workbook, read the formatted string of a cell using DisplayStringValue, obtain the workbook's CultureInfo currency symbol (or fall back to "$"), replace the default symbol, write the localized text to another cell, and save the file. Demonstrates currency localization with Aspose.Cells.
// Keywords: Aspose.Cells DisplayStringValue | C# get formatted cell value | localize currency symbol Aspose.Cells | Workbook.Settings.CultureInfo | replace $ with culture currency | Excel currency localization .NET | Aspose.Cells formatted string
// Common Searches: How to read a cell's displayed text in Aspose.Cells .NET | Replace $ sign with workbook culture currency symbol | Localize currency symbols in Excel using Aspose.Cells | Get formatted value of a cell and change currency symbol | Aspose.Cells currency localization example
// Developer Intent: Read a cell's formatted text, detect the default currency symbol, and substitute it with the symbol defined by the workbook’s CultureInfo (or a default) using Aspose.Cells for .NET.
// Use Cases: Display the correct localized currency in reports generated from Excel files. | Convert legacy workbooks that use a generic "$" symbol to region‑specific symbols. | Validate and store the localized string in another cell for downstream processing.
// AI Prompts: Show C# code that uses Aspose.Cells to get a cell's DisplayStringValue and replace the '$' with the workbook's CultureInfo currency symbol. | Write a reusable method that accepts a Worksheet and cell address, returns the formatted value with a localized currency symbol, and handles missing CultureInfo gracefully. | Explain how to fallback to a default currency symbol when Workbook.Settings.CultureInfo is not set in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;

// Load an Excel workbook, read the formatted string of a cell using DisplayStringValue, obtain the workbook's CultureInfo currency symbol (or fall back to "$"), replace the default symbol, write the localized text to another cell, and save the file. Demonstrates currency localization with Aspose.Cells.
class CurrencyLocalizationDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet and the target cell
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // Get the formatted string as shown in Excel
        string formattedValue = cell.DisplayStringValue;

        // Determine the currency symbol for the workbook's culture
        string cultureCurrencySymbol = workbook.Settings.CultureInfo != null
            ? workbook.Settings.CultureInfo.NumberFormat.CurrencySymbol
            : "$";

        // Replace the default "$" symbol with the localized currency symbol
        string localizedValue = formattedValue.Replace("$", cultureCurrencySymbol);

        // Output the results
        Console.WriteLine("Original formatted value: " + formattedValue);
        Console.WriteLine("Localized formatted value: " + localizedValue);

        // Optionally write the localized string to another cell for verification
        worksheet.Cells["B1"].PutValue(localizedValue);

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
