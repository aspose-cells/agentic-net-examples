// Title: Set Workbook Culture for Locale‑Specific Number Formatting with Aspose.Cells in C#
// Description: Shows how to assign a CultureInfo (e.g., German de‑DE) to a Workbook's Settings, optionally override decimal and group separators, apply a custom number format, and save the file, enabling locale‑aware number formatting during export.
// Keywords: Aspose.Cells | C# | workbook culture | CultureInfo | locale number formatting | de-DE | German number format | decimal separator | group separator | custom number format | Excel export
// Common Searches: Aspose.Cells set workbook culture .NET | C# change number format locale in Excel export | How to use German culture with Aspose.Cells | Set decimal separator in Aspose.Cells workbook | CultureInfo for number formatting Aspose.Cells
// Developer Intent: Configure the workbook’s culture to control how numbers (and optional properties) are formatted when the file is saved.
// Use Cases: Generate German‑locale financial reports where commas act as decimal separators and periods as thousand separators. | Create multi‑regional Excel exports that automatically adapt number formatting based on the user’s locale. | Override default separators for legacy systems that require specific symbols. | Apply a culture‑driven custom style to many cells without setting each format individually.
// AI Prompts: Write C# code that sets the workbook culture to French (fr‑FR) and formats dates and numbers accordingly using Aspose.Cells. | Explain how to detect a user's locale at runtime and switch wb.Settings.CultureInfo in Aspose.Cells. | Show how to export the same workbook in three locales (en‑US, de‑DE, ja‑JP) with appropriate number and date formats.

using System;
using System.Globalization;
using Aspose.Cells;

// Shows how to assign a CultureInfo (e.g., German de‑DE) to a Workbook's Settings, optionally override decimal and group separators, apply a custom number format, and save the file, enabling locale‑aware number formatting during export.
class ConfigureWorkbookCulture
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook to use a specific culture (German - Germany)
        wb.Settings.CultureInfo = new CultureInfo("de-DE");

        // Optional: explicitly set decimal and group separators (they follow the culture)
        wb.Settings.NumberDecimalSeparator = ',';
        wb.Settings.NumberGroupSeparator = '.';

        // Add sample data to demonstrate culture-aware formatting
        Worksheet sheet = wb.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(12345.67); // numeric value

        // Apply a custom number format; the separators will reflect the set culture
        Style style = wb.CreateStyle();
        style.Custom = "#,##0.00";
        cell.SetStyle(style);

        // Save the workbook
        wb.Save("CultureConfiguredWorkbook.xlsx");
    }
}
