using System;
using Aspose.Cells;

class DirectFormulaCalculationDemo
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // 1. Configure load options to skip parsing formulas when the workbook is opened
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false; // formulas remain as raw strings

        // 2. Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // 3. Get the first worksheet (the context for formula calculation)
        Worksheet sheet = workbook.Worksheets[0];

        // 4. Directly calculate a formula without placing it in a cell
        //    The formula will be evaluated as if it were entered in cell A1 of the worksheet
        string formula = "=SUM(A1:A3)";
        object result = sheet.CalculateFormula(formula);

        // 5. Display the calculated result
        Console.WriteLine($"Result of {formula} = {result}");

        // --------------------------------------------------------------
        // Optional: Parse all formulas that were not parsed during load
        // --------------------------------------------------------------
        // This step is useful if you later need the formulas to be fully parsed
        // (e.g., for further calculations or exporting).
        workbook.ParseFormulas(false);

        // After parsing, calculate all formulas in the workbook in the usual way
        workbook.CalculateFormula();

        // Save the workbook (optional)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}