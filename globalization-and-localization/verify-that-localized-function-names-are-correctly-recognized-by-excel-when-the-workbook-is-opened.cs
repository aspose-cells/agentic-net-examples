// Title: Verify that German localized Excel function SUMME is recognized and evaluated correctly with Aspose.Cells for .NET
// Common Searches: c# aspose.cells set workbook culture to de-de for german formula names | how to use SUMME function in an Aspose.Cells workbook | verify localized Excel formula persists after saving and loading with Aspose.Cells | calculate german localized functions in Aspose.Cells .NET example | check if Aspose.Cells preserves =SUMME(A1:A2) formula text
// Tags: set workbook culture Aspose.Cells | evaluate german localized function Aspose.Cells | save workbook with localized formula | load workbook preserving localized formula | c# aspose.cells formula localization verification

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets its culture to German (de-DE), writes sample data, applies the German SUMME formula, calculates and saves the file, then reloads it with the same culture. It confirms the formula text remains "=SUMME(A1:A2)" and the computed result is 30, demonstrating correct recognition of localized function names in Excel via Aspose.Cells for .NET.
class LocalizedFunctionVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set workbook culture to German so German function names are recognized
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the German localized function name for SUM: "SUMME"
            sheet.Cells["A3"].Formula = "SUMME(A1:A2)";

            // Calculate the formula in the current workbook
            workbook.CalculateFormula();

            // Save the workbook
            string filePath = "LocalizedFunction.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Load the workbook back
            Workbook loadedWorkbook = new Workbook(filePath);

            // Set the same culture for the loaded workbook so the localized formula can be evaluated
            loadedWorkbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Verify that the formula text remains the localized name (Aspose adds a leading '=')
            string loadedFormula = loadedWorkbook.Worksheets[0].Cells["A3"].Formula;
            Console.WriteLine("Loaded formula: " + loadedFormula); // Expected: =SUMME(A1:A2)

            // Evaluate the formula value after loading
            loadedWorkbook.CalculateFormula();
            double result = loadedWorkbook.Worksheets[0].Cells["A3"].DoubleValue;
            Console.WriteLine("Calculated result: " + result); // Expected: 30

            // Indicate verification success
            if (loadedFormula.Equals("=SUMME(A1:A2)", StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(result - 30) < 0.0001)
            {
                Console.WriteLine("Localized function name recognized correctly by Excel.");
            }
            else
            {
                Console.WriteLine("Verification failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
