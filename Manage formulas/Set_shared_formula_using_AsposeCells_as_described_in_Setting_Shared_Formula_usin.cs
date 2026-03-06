using Aspose.Cells;
using System;

class SetSharedFormulaOnLoad
{
    static void Main()
    {
        // Load options – disable formula parsing on open (optional)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set a shared formula in column B starting at B1 for 10 rows
        // Each B cell will calculate =A(row)*2
        cells["B1"].SetSharedFormula("=A1*2", 10, 1);

        // Recalculate all formulas after setting the shared formula
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}