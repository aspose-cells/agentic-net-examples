using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook with formula parsing disabled on open.
        // This ensures that existing formulas are not evaluated before we set our named range.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Get a reference to the first worksheet (used for building the formula string).
        Worksheet sheet = workbook.Worksheets[0];

        // Add a new global named range called "MyComplexFormula".
        int nameIdx = workbook.Worksheets.Names.Add("MyComplexFormula");
        Name namedRange = workbook.Worksheets.Names[nameIdx];

        // Assign a complex formula to the named range.
        // Example formula: =SUM(Sheet1!$A$1:$A$10) * AVERAGE(Sheet1!$B$1:$B$10)
        // The formula must start with an equal sign.
        namedRange.RefersTo = $"=SUM({sheet.Name}!$A$1:$A$10)*AVERAGE({sheet.Name}!$B$1:$B$10)";

        // Optionally calculate all formulas so that the named range can be used immediately.
        workbook.CalculateFormula();

        // Save the modified workbook.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}