using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options and disable formula parsing on open
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the XLSX workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and cell A1 (assumed to contain an IF formula)
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // Display the raw formula and the value (parsing was skipped, so value may be null)
        Console.WriteLine("Formula in A1: " + cell.Formula);
        Console.WriteLine("Raw value in A1 (parsing disabled): " + (cell.Value ?? "null"));

        // Reload the workbook with formula parsing enabled to get the evaluated result
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbookParsed = new Workbook("input.xlsx", loadOptions);
        Cell parsedCell = workbookParsed.Worksheets[0].Cells["A1"];

        // Display the evaluated value of the IF formula
        Console.WriteLine("Evaluated value in A1 (parsing enabled): " + parsedCell.Value);
    }
}