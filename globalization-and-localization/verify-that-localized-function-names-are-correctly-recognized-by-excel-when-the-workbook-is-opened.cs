// Title: Validate German localized Excel function SUMME recognition in Aspose.Cells for .NET by configuring workbook CultureInfo
// AI Prompts: Generate C# code that sets Workbook.Settings.CultureInfo to de-DE, inserts a SUMME(A1:A3) formula, calculates it, and verifies the result equals 60. | Write a .NET program that creates a workbook, populates cells A1‑A3 with numbers, applies the German function name SUMME in a formula, runs CalculateFormula, and checks the computed value. | Show how to programmatically confirm that a localized Excel function is evaluated correctly after setting the workbook culture in Aspose.Cells.
// Common Searches: how to configure Aspose.Cells workbook culture to German for localized formulas in C# | Aspose.Cells verify that SUMME formula returns correct result | C# example for using German function names like SUMME with Aspose.Cells | set workbook Settings.CultureInfo de-DE and calculate localized Excel functions in .NET
// Tags: Workbook.Settings.CultureInfo de-DE Aspose.Cells | German localized function SUMME formula | verify localized formula result C# | calculate Excel formulas with culture settings

using System;
using System.Globalization;
using Aspose.Cells;

// Creates a new workbook, sets its CultureInfo to German (de-DE), fills cells A1‑A3 with numeric values, assigns a formula using the German function name SUMME, calculates the formula, validates that the result is 60, and saves the workbook as LocalizedFunctionVerification.xlsx.
class LocalizedFunctionVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook culture to German to use German function names (e.g., SUMME)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");
            // Note: In newer Aspose.Cells versions the UseFormulaLocale property may be unavailable.
            // The CultureInfo setting is sufficient for this example.

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill cells A1, A2, A3 with numeric values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Set a formula using the German localized function name "SUMME"
            sheet.Cells["A4"].Formula = "SUMME(A1:A3)";

            // Calculate the formula (respects the set culture)
            workbook.CalculateFormula();

            // Verify that the calculated value matches the expected sum (60)
            double result = sheet.Cells["A4"].DoubleValue;
            if (Math.Abs(result - 60) < 0.0001)
            {
                Console.WriteLine("Localized function name recognized correctly. Result: " + result);
            }
            else
            {
                Console.WriteLine("Verification failed. Result: " + result);
            }

            // Save the workbook
            string outputPath = "LocalizedFunctionVerification.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
