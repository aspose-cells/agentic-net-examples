using System;
using Aspose.Cells;

class LoadWorkbookWithCustomOptions
{
    static void Main()
    {
        // Create LoadOptions and configure them to preserve formula calculations accurately
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true;                 // Parse formulas when the file is opened
        loadOptions.PreservePaddingSpacesInFormula = true;      // Keep original spaces/line‑breaks in formulas
        loadOptions.KeepUnparsedData = false;                   // Unnecessary unparsed data are discarded for better performance

        // Load the workbook using the custom LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Force calculation of all formulas to ensure values are up‑to‑date
        workbook.CalculateFormula();

        // Save the workbook with the calculated results
        workbook.Save("output.xlsx");
    }
}