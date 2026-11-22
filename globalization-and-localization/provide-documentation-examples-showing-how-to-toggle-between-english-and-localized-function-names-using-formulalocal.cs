// Title: How to toggle Excel function language between English and French using FormulaLocal in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook culture to French and use SOMME formula | switch between localized and English formulas with FormulaLocal in .NET | calculate Excel formulas in different languages using Aspose.Cells CultureInfo | use FormulaLocal to apply French SUM function in a .NET workbook | change workbook Settings.CultureInfo at runtime for multilingual formulas Aspose.Cells
// Tags: FormulaLocal with French SOMME function | Aspose.Cells culture info for multilingual formulas | C# workbook Settings.CultureInfo localization | switching Excel function language at runtime | localized Excel formulas in .NET using Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;

// The sample creates a workbook, fills cells B1:B5 with numbers, sets the workbook culture to French and assigns the French‑named SUM formula (SOMME) via FormulaLocal, calculates and prints the result, then switches the culture back to the invariant (English) culture, assigns the English SUM formula via FormulaLocal, recalculates, prints the result again, and finally saves the workbook as ToggleFormulaLocal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells B1:B5 with sample numbers 1 to 5
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 1].PutValue(i + 1); // Column index 1 = column B
        }

        // -------------------------------------------------
        // Use French localized function name via FormulaLocal
        // -------------------------------------------------
        // Set workbook culture to French (France)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Assign a formula using the French name for SUM ("SOMME")
        sheet.Cells["A1"].FormulaLocal = "SOMME(B1:B5)";

        // Calculate formulas
        workbook.CalculateFormula();

        // Output the result (should be 15)
        Console.WriteLine("Result with French function (SOMME): " + sheet.Cells["A1"].StringValue);

        // -------------------------------------------------
        // Switch back to English (invariant culture) and use English name
        // -------------------------------------------------
        workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

        // Assign a formula using the English name for SUM
        sheet.Cells["A2"].FormulaLocal = "SUM(B1:B5)";

        // Recalculate formulas
        workbook.CalculateFormula();

        // Output the result (should also be 15)
        Console.WriteLine("Result with English function (SUM): " + sheet.Cells["A2"].StringValue);

        // Save the workbook (optional)
        workbook.Save("ToggleFormulaLocal.xlsx");
    }
}
