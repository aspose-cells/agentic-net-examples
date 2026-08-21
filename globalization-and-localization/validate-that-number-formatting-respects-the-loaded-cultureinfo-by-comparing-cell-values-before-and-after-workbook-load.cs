// Title: Validate CultureInfo‑specific number formatting in Aspose.Cells (C#)
// Description: Creates a workbook, writes a numeric value with the built‑in "0.00" format, saves it, reads the cell's string representation using the default culture, then reloads the file with a German CultureInfo (de‑DE) and compares the string values to confirm that the decimal separator follows the loaded locale.
// Keywords: Aspose.Cells | CultureInfo | C# | number formatting | decimal separator | LoadOptions | de-DE | globalization | localization | XLSX validation | unit test
// Common Searches: Aspose.Cells verify CultureInfo on load | C# check decimal separator after loading workbook | how to test number format with different locales in Aspose.Cells | load XLSX with German culture Aspose.Cells
// Developer Intent: Confirm that a workbook loaded with a specific CultureInfo renders numeric cells using that locale’s decimal separator.
// Use Cases: Automated test to ensure financial reports display commas as decimal points when opened with a German locale. | Quality‑gate check that saved spreadsheets retain locale‑aware formatting after being transferred between regions. | Integration scenario where a report generated in one language is reviewed in another, requiring verification of correct number representation.
// AI Prompts: Generate an xUnit test that creates a workbook, applies the built‑in number format, saves it, reloads with a given CultureInfo, and asserts the expected string value. | Write a reusable method that loads an XLSX file with LoadOptions.CultureInfo and returns the formatted string of a specified cell. | Explain how Aspose.Cells selects the decimal separator for built‑in number formats when a CultureInfo is supplied in LoadOptions.

using System;
using System.Globalization;
using Aspose.Cells;

// Creates a workbook, writes a numeric value with the built‑in "0.00" format, saves it, reads the cell's string representation using the default culture, then reloads the file with a German CultureInfo (de‑DE) and compares the string values to confirm that the decimal separator follows the loaded locale.
class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // Create a workbook, put a numeric value and apply a built‑in number format.
        // -------------------------------------------------
        Workbook wb = new Workbook();                         // create workbook
        Worksheet ws = wb.Worksheets[0];
        Cell cell = ws.Cells["A1"];
        cell.PutValue(1234.56);                               // numeric value

        // Apply built‑in format "0.00" (Number = 2) which uses the current culture's decimal separator.
        Style style = wb.CreateStyle();
        style.Number = 2;                                     // decimal with two places
        cell.SetStyle(style);

        // Save the workbook to a temporary file.
        string filePath = "culture_test.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // Capture the string representation using the default (environment) culture.
        string beforeLoad = cell.StringValue;
        Console.WriteLine($"Before load (default culture): {beforeLoad}");

        // -------------------------------------------------
        // Load the same file with a different CultureInfo (German uses comma as decimal separator).
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("de-DE");   // set culture for loading

        Workbook loadedWb = new Workbook(filePath, loadOptions); // load with culture
        Worksheet loadedWs = loadedWb.Worksheets[0];
        Cell loadedCell = loadedWs.Cells["A1"];

        // Get the string representation after loading with German culture.
        string afterLoad = loadedCell.StringValue;
        Console.WriteLine($"After load (de-DE culture): {afterLoad}");

        // -------------------------------------------------
        // Compare the two representations to verify culture‑aware formatting.
        // -------------------------------------------------
        bool formattingUnchanged = beforeLoad == afterLoad;
        Console.WriteLine($"Formatting unchanged: {formattingUnchanged}");
    }
}
