// Title: Validate Named Range References After German Localization in Aspose.Cells for .NET
// Description: Shows how to set the workbook region to Germany, map the SUM function to the German name SUMME, create a named range, retrieve its localized RefersTo string, apply a FormulaLocal using SUMME, calculate the result, and confirm the named range still points to the original cells before saving.
// Keywords: Aspose.Cells | German localization | named range | RefersTo | FormulaLocal | SUMME | .NET | globalization settings | SettableGlobalizationSettings | region Germany
// Common Searches: Aspose.Cells German locale named range | keep named range address after localization | FormulaLocal SUMME example C# | GetRefersTo localized string Aspose.Cells | set workbook region Germany Aspose.Cells
// Developer Intent: Verify that a named range continues to reference the correct cells when the workbook is localized to German and when using a localized formula.
// Use Cases: Call GetRefersTo(false, true) on a Name object to ensure the A1 address remains unchanged under German settings. | Assign a FormulaLocal string using the German function name SUMME that references the named range and validate that the calculated value equals the sum of the range. | Retrieve the Range via Name.GetRange(), inspect its RefersTo property and the first cell value to confirm globalization did not alter the range. | Save the workbook and open it in Excel to see the German function name displayed while the range stays intact.
// AI Prompts: Generate C# code that creates a named range, applies German localization with SettableGlobalizationSettings, and checks that the RefersTo address stays the same. | Explain step‑by‑step how FormulaLocal works with localized function names and named ranges in Aspose.Cells for .NET. | Provide a unit‑test example that asserts the RefersTo string of a named range is unchanged after setting workbook.Settings.Region to Germany.

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangeLocalizationDemo
{
    // Shows how to set the workbook region to Germany, map the SUM function to the German name SUMME, create a named range, retrieve its localized RefersTo string, apply a FormulaLocal using SUMME, calculate the result, and confirm the named range still points to the original cells before saving.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Set the workbook region to Germany to enable German localization
                workbook.Settings.Region = CountryCode.Germany;

                // Create custom globalization settings and map the standard SUM function to German "SUMME"
                SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();
                gSettings.SetLocalFunctionName("SUM", "SUMME", true);
                workbook.Settings.GlobalizationSettings = gSettings;

                // Populate some data that will be used by the named range
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a named range that refers to the three cells above (standard A1 notation)
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];
                myRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Retrieve the RefersTo string in localized (German) format
                // isR1C1 = false (A1 notation), isLocal = true (apply locale)
                string localizedRefersTo = myRange.GetRefersTo(false, true);
                Console.WriteLine("Localized RefersTo: " + localizedRefersTo);
                // Expected output: "=Sheet1!$A$1:$A$3" (same address, but locale flag is honored)

                // Use the localized function name in a formula that references the named range
                Cell formulaCell = sheet.Cells["B1"];
                formulaCell.FormulaLocal = "=SUMME(MyRange)";

                // Calculate the workbook to evaluate the formula
                workbook.CalculateFormula();

                // Display the result of the localized formula
                Console.WriteLine("Result of localized formula (SUMME(MyRange)): " + formulaCell.Value);

                // Verify that the named range still points to the correct range after localization
                Aspose.Cells.Range range = myRange.GetRange();
                Console.WriteLine($"Named range address after localization: {range.RefersTo}");
                Console.WriteLine($"First cell value in range: {range[0, 0].Value}");

                // Save the workbook
                string outputPath = "NamedRangeLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
