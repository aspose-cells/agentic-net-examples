using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure load options to parse formulas when the workbook is opened.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true; // Enables automatic formula parsing during load.

        // Load the XLSX file with the specified options.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Evaluate all formulas in the loaded workbook.
        workbook.CalculateFormula();

        // Save the workbook so that the calculated values are persisted.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}