using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook. ParsingFormulaOnOpen ensures formulas are parsed during load.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Parse any formulas that might not have been parsed when the file was loaded.
        workbook.ParseFormulas(false);

        // Calculate all formulas in the workbook, including the IFNA function.
        workbook.CalculateFormula();

        // Retrieve and display the result of a cell that contains the IFNA formula.
        // Adjust the cell reference as needed for your specific workbook.
        object ifnaResult = workbook.Worksheets[0].Cells["B2"].Value;
        Console.WriteLine("IFNA result: " + ifnaResult);

        // Save the workbook after calculation.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}