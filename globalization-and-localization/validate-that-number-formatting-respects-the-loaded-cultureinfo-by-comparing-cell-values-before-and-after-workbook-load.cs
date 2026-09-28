// Title: Validate that Aspose.Cells preserves numeric value while applying en-US and fr-FR CultureInfo formatting before and after workbook load in C#
// AI Prompts: Generate C# code that creates a workbook, sets Settings.CultureInfo to en-US, writes a double to a cell with built‑in number format 2, saves to a MemoryStream, reloads the stream with Settings.CultureInfo set to fr-FR, and reads both the raw double and the cell's StringValue. | Add logic that asserts the raw double values before and after loading are equal within a tolerance and that the formatted string after loading contains a comma as the decimal separator. | Convert the example into an xUnit test method that fails when the locale‑specific formatted string does not match the expected French pattern while still confirming value equality.
// Common Searches: Aspose.Cells check numeric cell value remains same after changing workbook CultureInfo | C# how to compare Excel cell raw value and displayed string with different locales using Aspose | verify that number format uses comma decimal separator after loading workbook with fr-FR in Aspose.Cells | unit test for culture‑specific number formatting in Aspose.Cells .NET
// Tags: Aspose.Cells cultureinfo numeric formatting validation | C# verify raw double vs formatted string Aspose.Cells | load workbook with different locale Aspose.Cells | Excel number format locale test .NET | built‑in number format 0.00 Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook with en-US CultureInfo, writes 1234.56 to cell A1 using the built‑in "0.00" format, saves it to a memory stream, reloads the workbook with fr-FR CultureInfo, and then reads the cell's raw double value and its formatted string. It demonstrates that the underlying numeric value stays identical while the displayed string reflects the French locale formatting.
class CultureInfoValidation
{
    static void Main()
    {
        // Create a new workbook and set its CultureInfo to en-US
        Workbook wb = new Workbook();
        wb.Settings.CultureInfo = new CultureInfo("en-US");

        // Access the first worksheet and cell A1
        Worksheet sheet = wb.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Set a numeric value and apply a built‑in number format with two decimals
        cell.PutValue(1234.56);
        Style style = cell.GetStyle();
        style.Number = 2; // Built‑in format: "0.00"
        cell.SetStyle(style);

        // Capture the value and formatted string before saving
        double originalValue = cell.Value is double d ? d : 0.0;
        string originalFormatted = cell.StringValue;

        // Save the workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            wb.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;

            // Load the workbook from the stream and set CultureInfo to fr-FR
            Workbook loadedWb = new Workbook(ms);
            loadedWb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Access the same cell in the loaded workbook
            Worksheet loadedSheet = loadedWb.Worksheets[0];
            Cell loadedCell = loadedSheet.Cells["A1"];

            // Capture the value and formatted string after loading
            double loadedValue = loadedCell.Value is double ld ? ld : 0.0;
            string loadedFormatted = loadedCell.StringValue;

            // Validate that the numeric values are equal
            bool valuesEqual = Math.Abs(originalValue - loadedValue) < 1e-10;

            // Output the results
            Console.WriteLine($"Original Value: {originalValue}");
            Console.WriteLine($"Loaded Value:   {loadedValue}");
            Console.WriteLine($"Values equal:   {valuesEqual}");
            Console.WriteLine($"Original formatted (en-US): {originalFormatted}");
            Console.WriteLine($"Loaded formatted (fr-FR):   {loadedFormatted}");
        }
    }
}
