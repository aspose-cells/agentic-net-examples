using System;
using System.Globalization;
using Aspose.Cells;

class FormulaLocalDemo
{
    static void Main()
    {
        // Path to the source workbook (must exist)
        string inputPath = "input.xlsx";

        // Path for the resulting workbook
        string outputPath = "output.xlsx";

        // Create LoadOptions and set the culture to German (de-DE)
        // This influences number formatting and formula parsing on load
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("de-DE");

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Set the workbook's regional setting to Germany.
        // This affects how FormulaLocal is rendered.
        workbook.Settings.Region = CountryCode.Germany;

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Prepare some numeric data for formulas
        cells["B1"].PutValue(10);
        cells["C1"].PutValue(20);

        // -------------------------------------------------
        // 1. Set a formula using the standard (English) name
        // -------------------------------------------------
        Cell cellA1 = cells["A1"];
        cellA1.Formula = "=SUM(B1:C1)";

        // Show both the standard and the localized (German) representation
        Console.WriteLine("Standard Formula (A1): " + cellA1.Formula);
        Console.WriteLine("Localized Formula (A1): " + cellA1.FormulaLocal);

        // -------------------------------------------------
        // 2. Set a formula using the German localized name via FormulaLocal
        // -------------------------------------------------
        Cell cellA2 = cells["A2"];
        cellA2.FormulaLocal = "=SUMME(B1:C1)";

        // After assigning the localized formula, display both forms
        Console.WriteLine("\nAfter setting FormulaLocal (A2):");
        Console.WriteLine("Standard Formula (A2): " + cellA2.Formula);
        Console.WriteLine("Localized Formula (A2): " + cellA2.FormulaLocal);

        // -------------------------------------------------
        // 3. Calculate formulas to obtain results
        // -------------------------------------------------
        workbook.CalculateFormula();

        Console.WriteLine("\nCalculated Values:");
        Console.WriteLine("A1 (SUM): " + cellA1.Value);
        Console.WriteLine("A2 (SUMME): " + cellA2.Value);

        // -------------------------------------------------
        // 4. Save the workbook with the changes
        // -------------------------------------------------
        workbook.Save(outputPath);
    }
}