using System;
using Aspose.Cells;

class ApplySharedFormulaOnLoad
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Configure load options to skip formula parsing on open.
        // This allows us to set a shared formula after loading without immediate parsing.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the workbook with the specified options.
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Example: Apply a shared formula in column B that multiplies column A values by 2.
        // The shared formula will be applied to 10 rows starting from B1.
        cells["B1"].SetSharedFormula("=A1*2", 10, 1);

        // Parse all formulas that were not parsed during load.
        // 'false' means do not ignore errors; an exception will be thrown if a formula is invalid.
        workbook.ParseFormulas(false);

        // Calculate the workbook to evaluate the newly set shared formulas.
        workbook.CalculateFormula();

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}