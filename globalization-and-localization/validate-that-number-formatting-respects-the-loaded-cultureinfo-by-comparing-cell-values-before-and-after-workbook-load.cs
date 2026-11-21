// Title: Check that a custom number format respects a specified CultureInfo when saving and reloading an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify that number format stays the same after loading a workbook with CultureInfo | how to preserve localized numeric formatting when saving and reopening an .xlsx with Aspose.Cells | using LoadOptions to maintain French decimal separator in Excel file via Aspose.Cells | compare cell.StringValue before and after workbook reload to test culture-specific format | C# Aspose.Cells memory stream save load culture-specific number format validation
// Tags: custom number format cultureinfo | Aspose.Cells loadoptions cultureinfo | excel localization number formatting c# | memorystream workbook save load Aspose.Cells | validate number format preservation

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example sets the workbook's CultureInfo to French, writes a numeric value with a custom "#,##0.00" format, captures the formatted string, saves the workbook to a MemoryStream, reloads it using LoadOptions with the same CultureInfo, retrieves the formatted string again, and compares the two strings to confirm that the culture‑specific formatting is preserved.
class Program
{
    static void Main()
    {
        // Define the culture to test (e.g., French uses comma as decimal separator)
        CultureInfo testCulture = new CultureInfo("fr-FR");

        // ---------- Create workbook ----------
        Workbook wb = new Workbook();
        // Apply the culture to the workbook settings
        wb.Settings.CultureInfo = testCulture;

        // Write a numeric value and apply a custom number format
        Worksheet ws = wb.Worksheets[0];
        Cell cell = ws.Cells["A1"];
        double numericValue = 1234.56;
        cell.PutValue(numericValue);

        // Use a format that displays the decimal separator according to the culture
        Style style = cell.GetStyle();
        style.Custom = "#,##0.00";
        cell.SetStyle(style);

        // Capture the formatted string before saving
        string formattedBeforeSave = cell.StringValue;

        // ---------- Save workbook to memory ----------
        using (MemoryStream ms = new MemoryStream())
        {
            wb.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream position for reading

            // ---------- Load workbook with the same culture ----------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = testCulture
            };
            Workbook loadedWb = new Workbook(ms, loadOptions);
            Cell loadedCell = loadedWb.Worksheets[0].Cells["A1"];
            string formattedAfterLoad = loadedCell.StringValue;

            // ---------- Validation ----------
            bool formattingPreserved = formattedBeforeSave == formattedAfterLoad;
            Console.WriteLine($"Formatted before save: {formattedBeforeSave}");
            Console.WriteLine($"Formatted after load : {formattedAfterLoad}");
            Console.WriteLine($"Formatting preserved : {formattingPreserved}");
        }
    }
}
