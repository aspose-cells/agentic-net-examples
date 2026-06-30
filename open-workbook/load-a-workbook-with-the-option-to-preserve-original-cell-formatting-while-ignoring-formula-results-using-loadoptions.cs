using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – load workbook preserving formatting and ignoring formula results
class Program
{
    static void Main()
    {
        // Configure load options
        LoadOptions loadOptions = new LoadOptions
        {
            // Do not evaluate formulas when opening the file; keep the original formula text.
            ParsingFormulaOnOpen = false
            // Cell formatting is preserved by default; no additional setting required.
        };

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example usage: output the number of worksheets to verify successful load
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);

        // Save the workbook (format and formatting are retained)
        workbook.Save("output.xlsx");
    }
}