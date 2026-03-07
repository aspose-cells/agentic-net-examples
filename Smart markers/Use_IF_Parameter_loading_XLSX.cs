using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Example parameter that decides whether formulas should be parsed on load
        bool parseFormulas = false; // Change to true to enable parsing

        // Create load options (uses LoadOptions() constructor)
        LoadOptions loadOptions = new LoadOptions();

        // Set the ParsingFormulaOnOpen property based on the parameter (uses property setter)
        if (parseFormulas)
        {
            loadOptions.ParsingFormulaOnOpen = true;
        }
        else
        {
            loadOptions.ParsingFormulaOnOpen = false;
        }

        // Load the XLSX workbook with the specified load options (uses Workbook(string, LoadOptions) constructor)
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and a sample cell
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // Output the formula and value of the cell to demonstrate the effect of the option
        Console.WriteLine("Cell A1 Formula: " + cell.Formula);
        Console.WriteLine("Cell A1 Value: " + cell.Value);

        // Save the workbook (uses Workbook.Save(string) method)
        workbook.Save("output.xlsx");
    }
}