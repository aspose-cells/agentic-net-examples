// Title: Validate that Aspose.Cells preserves CultureInfo number formatting before and after saving to XLSX
// Common Searches: how to test number format culture in Aspose.Cells C# | verify French numeric formatting after saving Excel with Aspose.Cells | compare cell stringvalue before and after workbook load Aspose.Cells .NET | preserve CultureInfo formatting when reloading workbook using Aspose.Cells | C# Aspose.Cells number format 0.00 culture-specific validation
// Tags: Aspose.Cells cultureinfo number formatting | validate number format after workbook save load | French locale numeric format Excel .NET | built‑in number format 0.00 Aspose.Cells | XLSX culture‑aware formatting verification

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets Settings.CultureInfo to French (fr-FR), writes a double to cell A1, applies the built‑in 0.00 number format, captures the formatted string, saves the workbook to a memory stream as XLSX, reloads it, reassigns the same CultureInfo, captures the formatted string again, and compares the two values to confirm that number formatting respects the CultureInfo across save and load operations.
class Program
{
    static void Main()
    {
        // Create a new workbook and set its CultureInfo to French (France)
        var workbook = new Workbook();
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        double value = 1234.56;
        sheet.Cells["A1"].PutValue(value);

        // Apply a number format with two decimal places
        var style = sheet.Cells["A1"].GetStyle();
        style.Number = 2; // Built‑in format: 0.00
        sheet.Cells["A1"].SetStyle(style);

        // Capture the formatted string before saving
        string formattedBefore = sheet.Cells["A1"].StringValue;

        // Save the workbook to a memory stream (using the provided save rule)
        using (var stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0;

            // Load the workbook from the stream (using the provided load rule)
            var loadedWorkbook = new Workbook(stream);
            loadedWorkbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            var loadedSheet = loadedWorkbook.Worksheets[0];
            string formattedAfter = loadedSheet.Cells["A1"].StringValue;

            // Compare the formatted strings to validate culture‑aware formatting
            bool formattingRespectsCulture = formattedBefore == formattedAfter;

            Console.WriteLine($"Formatted before load: {formattedBefore}");
            Console.WriteLine($"Formatted after load:  {formattedAfter}");
            Console.WriteLine($"Formatting respects CultureInfo: {formattingRespectsCulture}");
        }
    }
}
