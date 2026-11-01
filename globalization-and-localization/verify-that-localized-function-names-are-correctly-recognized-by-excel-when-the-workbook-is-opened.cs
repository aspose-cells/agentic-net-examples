// Title: How to verify German localized Excel functions like SUMME are recognized with Aspose.Cells for .NET
// Common Searches: asp.net verify german Excel function SUMME using Aspose.Cells | set workbook culture to de-DE and recalculate formulas with Aspose.Cells C# | load workbook from memory stream and evaluate localized formulas Aspose.Cells | check if Aspose.Cells recognizes localized function names after reload | example of using Aspose.Cells to calculate German functions in .xlsx
// Tags: Aspose.Cells workbook cultureinfo de-de | German Excel function SUMME evaluation | formula recalculation after workbook reload | localized function name verification in XLSX | C# Aspose.Cells localized formula example

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, sets its CultureInfo to German (de-DE), writes numeric values, applies the German SUMME formula, saves to a memory stream, reloads the workbook, reapplies the German culture, recalculates formulas, and confirms that cell A3 contains the expected result of 30.
class LocalizedFunctionVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set culture to German to use German function names (e.g., SUMME for SUM)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the German localized function name for SUM (SUMME) in a formula
            sheet.Cells["A3"].Formula = "SUMME(A1:A2)";

            // Save the workbook to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Load the workbook back
                Workbook loadedWorkbook = new Workbook(ms);

                // Ensure the same culture is applied to the loaded workbook
                loadedWorkbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Recalculate formulas
                loadedWorkbook.CalculateFormula();

                // Retrieve the calculated value from cell A3
                double result = loadedWorkbook.Worksheets[0].Cells["A3"].DoubleValue;

                // Verify the result (10 + 20 = 30)
                Console.WriteLine("Calculated result in A3: " + result);
                Console.WriteLine("Verification " + (Math.Abs(result - 30) < 0.0001 ? "passed" : "failed"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
