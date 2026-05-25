using System;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set the workbook region to Germany to simulate a locale that uses
            // localized function names and semicolon argument separators.
            workbook.Settings.Region = CountryCode.Germany;

            // Populate sample data
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(15);

            // Assign a formula using the localized (German) function name and separator.
            // In German the SUM function is "SUMME" and arguments are separated by ';'.
            sheet.Cells["B1"].FormulaLocal = "=SUMME(A1;A2)";

            // Retrieve the standard (English) representation of the formula.
            // Aspose.Cells automatically converts the localized formula to the
            // standard one and stores it in the Formula property.
            string standardFormula = sheet.Cells["B1"].Formula; // e.g., "=SUM(A1,A2)"

            // Verify that the standard formula uses English function names and commas.
            bool usesEnglishFunction = standardFormula.IndexOf("SUM", StringComparison.OrdinalIgnoreCase) >= 0;
            bool usesCommaSeparator = standardFormula.Contains(",");

            Console.WriteLine("Standard Formula: " + standardFormula);
            Console.WriteLine("Uses English function name: " + usesEnglishFunction);
            Console.WriteLine("Uses comma as argument separator: " + usesCommaSeparator);

            // Optionally calculate the formula to ensure it works.
            workbook.CalculateFormula();

            Console.WriteLine("Calculated value in B1: " + sheet.Cells["B1"].Value);

            // Save the workbook (lifecycle rule)
            workbook.Save("FormulaVerificationResult.xlsx");
        }
    }
}