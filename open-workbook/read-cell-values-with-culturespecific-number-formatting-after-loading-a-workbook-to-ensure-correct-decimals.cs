// Title: Read culture‑specific number formatting from cells after loading a workbook (Aspose.Cells for .NET)
// Description: Demonstrates how to load an Excel file with a specific CultureInfo (e.g., German de‑DE) using Aspose.Cells LoadOptions, adjust NumberDecimalSeparator and NumberGroupSeparator, and retrieve both the raw DoubleValue and the locale‑aware StringValue of a cell.
// Keywords: Aspose.Cells | C# | .NET | LoadOptions CultureInfo | de-DE | German number format | NumberDecimalSeparator | NumberGroupSeparator | cell DoubleValue | cell StringValue | locale aware Excel reading | read cell value culture specific
// Common Searches: Aspose.Cells load workbook with German culture | read cell value with comma decimal separator Aspose.Cells | LoadOptions.CultureInfo example C# | set number decimal separator after loading workbook | get formatted string from cell Aspose.Cells
// Developer Intent: Load a workbook with a designated CultureInfo and obtain cell values that respect the locale’s decimal and grouping symbols.
// Use Cases: Show financial numbers saved with dot decimals to users in Germany, displaying commas as decimal separators. | Perform calculations using the raw DoubleValue while presenting the StringValue formatted for the user’s locale in a UI. | Adjust workbook number settings after loading to match a target culture before exporting to PDF or other formats.
// AI Prompts: Generate C# code that loads an Excel file with French CultureInfo using Aspose.Cells, sets appropriate decimal and group separators, and reads both DoubleValue and StringValue from a cell. | Explain how LoadOptions.CultureInfo and Workbook.Settings.NumberDecimalSeparator interact to affect cell formatting in Aspose.Cells. | Provide a step‑by‑step guide to retrieve locale‑specific formatted strings from cells after loading a workbook with a custom CultureInfo.

using System;
using System.Globalization;
using Aspose.Cells;

// Demonstrates how to load an Excel file with a specific CultureInfo (e.g., German de‑DE) using Aspose.Cells LoadOptions, adjust NumberDecimalSeparator and NumberGroupSeparator, and retrieve both the raw DoubleValue and the locale‑aware StringValue of a cell.
class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample workbook and put a numeric value using dot as
        //    decimal separator (default invariant culture).
        // -----------------------------------------------------------------
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue(1234.56);
        string filePath = "sample.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // -----------------------------------------------------------------
        // 2. Load the workbook with a specific culture (German) that uses
        //    a comma as decimal separator.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("de-DE"); // German culture
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Optionally adjust workbook settings to match the culture.
        // -----------------------------------------------------------------
        loadedWb.Settings.NumberDecimalSeparator = ',';   // comma for decimals
        loadedWb.Settings.NumberGroupSeparator = '.';    // dot for thousands

        // -----------------------------------------------------------------
        // 4. Read the cell value.
        //    - DoubleValue gives the raw numeric value.
        //    - StringValue returns the formatted string according to the
        //      workbook's culture/number settings.
        // -----------------------------------------------------------------
        Cell cell = loadedWb.Worksheets[0].Cells["A1"];
        double numericValue = cell.DoubleValue;
        string formattedValue = cell.StringValue;

        Console.WriteLine($"Numeric value (double): {numericValue}");
        Console.WriteLine($"Formatted string (German culture): {formattedValue}");
    }
}
