// Title: Aspose.Cells for .NET – Apply German locale to named‑range formulas (FormulaLocal)
// Description: This example creates a workbook, sets its region to Germany, defines custom globalization settings that map English functions (SUM, AVERAGE) to German equivalents (SUMME, MITTELWERT), adds a named range "MyRange" (B1:B5), writes an English formula with `Formula` and a German formula with `FormulaLocal`, calculates both, prints the results, and saves the file. The workbook works with a German‑installed Excel and respects the system locale.
// Keywords: Aspose.Cells | German locale | .NET | C# | SettableGlobalizationSettings | FormulaLocal | named range localization | Excel German functions | SUMME | MITTELWERT | region Germany | function translation
// Common Searches: Aspose.Cells German locale example | How to localize Excel formulas to German in C# | Set workbook region to Germany Aspose.Cells | FormulaLocal German named range Aspose.Cells | Map English Excel functions to German using Aspose.Cells | German Excel function names SUMME Aspose.Cells
// Developer Intent: Generate a workbook that shows German‑localized formulas for a named range while keeping the workbook region set to Germany, enabling compatibility with German‑installed Excel.
// Use Cases: Create financial reports that open correctly in German Excel by translating function names. | Automated tests that verify English and German formulas return identical results for the same named range. | Produce multi‑language workbooks that switch between English and German locales at runtime. | Migrate existing English workbooks to German markets without manual formula editing.
// AI Prompts: Add additional German function mappings (e.g., MIN → MIN, MAX → MAX) using SettableGlobalizationSettings. | Show how to retrieve the localized formula string from a cell after calculation. | Demonstrate switching between English and German locales for the same workbook while preserving named ranges. | Explain how to configure the system locale for Aspose.Cells to match a German Excel installation. | Provide unit‑test code that asserts both English and German formulas produce the same value.

using System;
using Aspose.Cells;

namespace AsposeCellsGermanLocaleDemo
{
    // This example creates a workbook, sets its region to Germany, defines custom globalization settings that map English functions (SUM, AVERAGE) to German equivalents (SUMME, MITTELWERT), adds a named range "MyRange" (B1:B5), writes an English formula with `Formula` and a German formula with `FormulaLocal`, calculates both, prints the results, and saves the file. The workbook works with a German‑installed Excel and respects the system locale.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set the workbook region to Germany (German locale)
            workbook.Settings.Region = CountryCode.Germany;

            // Create custom globalization settings and map English function names to German equivalents
            SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();
            gSettings.SetLocalFunctionName("SUM", "SUMME", true);          // SUM → SUMME
            gSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true); // AVERAGE → MITTELWERT

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = gSettings;

            // Fill some sample data in column B (B1:B5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[$"B{i + 1}"].PutValue(i + 1); // Values 1,2,3,4,5
            }

            // Define a named range "MyRange" that refers to B1:B5
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "='Sheet1'!$B$1:$B$5";

            // Use the standard English formula with the named range
            Cell cellStd = sheet.Cells["A1"];
            cellStd.Formula = "=SUM(MyRange)";

            // Use the German localized formula with the same named range via FormulaLocal
            Cell cellLocal = sheet.Cells["A2"];
            cellLocal.FormulaLocal = "=SUMME(MyRange)";

            // Calculate all formulas
            workbook.CalculateFormula();

            // Output results to console
            Console.WriteLine($"Standard formula result (A1): {cellStd.Value}");
            Console.WriteLine($"German localized formula result (A2): {cellLocal.Value}");

            // Save the workbook
            workbook.Save("GermanLocaleNamedRangeDemo.xlsx");
        }
    }
}
