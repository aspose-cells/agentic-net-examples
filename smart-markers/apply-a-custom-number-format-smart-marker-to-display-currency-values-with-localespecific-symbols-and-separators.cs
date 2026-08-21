// Title: Apply a Locale‑Aware Custom Currency Format with Aspose.Cells for .NET
// Description: Demonstrates how to set a workbook's CultureInfo, create a custom accounting‑style number format, apply only the NumberFormat flag to a range, and save the Excel file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | custom number format | currency formatting | locale specific | CultureInfo | French currency | accounting format | SetCustom | StyleFlag | Excel export
// Common Searches: Aspose.Cells set workbook culture for currency | C# custom accounting number format Aspose.Cells | locale aware currency format Excel using Aspose | apply number format flag to range Aspose.Cells | French euro format in Aspose.Cells workbook
// Developer Intent: Generate an Excel workbook where numeric cells are displayed as currency using a custom format that automatically reflects the workbook’s locale settings.
// Use Cases: Create invoices for French clients with € symbols and proper thousand separators. | Build a multi‑regional financial report that adapts currency symbols by changing CultureInfo only. | Export .NET data to Excel while guaranteeing locale‑correct currency display without manual string manipulation.
// AI Prompts: Write C# code with Aspose.Cells to apply a custom accounting number format that inherits the workbook’s CultureInfo for currency symbols and separators. | Explain the steps to create a Style, set a custom pattern with SetCustom, enable only NumberFormat via StyleFlag, and apply it to a cell range. | Provide a test plan to verify that the French currency format appears correctly in the generated Excel file.

using System;
using System.Globalization;
using Aspose.Cells;

// Demonstrates how to set a workbook's CultureInfo, create a custom accounting‑style number format, apply only the NumberFormat flag to a range, and save the Excel file using Aspose.Cells in C#.
class CurrencySmartMarkerDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Set the workbook culture to French (France) – this changes the currency symbol and separators
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Get the first worksheet and its cells collection
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some numeric values that will be displayed as currency
            cells["A1"].PutValue(1234.56);
            cells["A2"].PutValue(7890.12);
            cells["A3"].PutValue(345.67);

            // Create a style with a custom accounting‑style currency format.
            // The pattern is invariant; the actual symbol and separators are taken from the workbook culture.
            Style style = wb.CreateStyle();
            string customPattern = "_-€ * #,##0.00_-;_-€ * -#,##0.00_-;_-€ * \"-\"??_-;_-@_-";
            style.SetCustom(customPattern, true); // true = prefer built‑in if it matches

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A3
            Aspose.Cells.Range range = cells.CreateRange("A1:A3");
            range.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "CurrencySmartMarkerDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
