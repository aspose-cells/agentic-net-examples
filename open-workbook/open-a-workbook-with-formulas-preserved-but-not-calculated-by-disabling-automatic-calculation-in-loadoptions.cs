// Title: Load an Excel workbook without evaluating formulas using Aspose.Cells LoadOptions (C#)
// Description: Demonstrates how to create a LoadOptions object, set ParsingFormulaOnOpen to false, and open a workbook so that formulas are loaded as raw text and are not calculated automatically. The example shows reading a formula, checking the uncomputed value, and optionally triggering manual calculation later.
// Keywords: Aspose.Cells LoadOptions | disable formula calculation | ParsingFormulaOnOpen false | open workbook without evaluating formulas | preserve formulas as text C# | manual formula calculation Aspose.Cells
// Common Searches: How to open Excel with formulas unchanged Aspose.Cells | Disable automatic formula evaluation on load .NET | Load workbook as raw formulas Aspose.Cells C# | Prevent formula calculation when opening workbook | Aspose.Cells load options for formula preservation
// Developer Intent: Open an Excel file while keeping all formulas intact and avoiding any automatic calculation.
// Use Cases: Inspect or edit formula strings without triggering calculations. | Export a workbook to another format while preserving original formulas. | Improve performance when loading large, formula‑heavy workbooks and calculate only needed cells later.
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, keeps formulas uncalculated, then calculates a specific range on demand. | Show how to iterate through all cells after loading with ParsingFormulaOnOpen false and retrieve each cell's raw formula. | Provide an example that saves the workbook after calling CalculateFormula, ensuring the original formulas remain unchanged.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLoadExample
{
    // Demonstrates how to create a LoadOptions object, set ParsingFormulaOnOpen to false, and open a workbook so that formulas are loaded as raw text and are not calculated automatically. The example shows reading a formula, checking the uncomputed value, and optionally triggering manual calculation later.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains formulas
            string inputFile = "input.xlsx";

            // Create LoadOptions and disable formula parsing on open
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false; // formulas will be loaded as raw strings, not calculated

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example cell that contains a formula (adjust address as needed)
            Cell formulaCell = sheet.Cells["A1"];

            // Display the formula text and the current value (should be default/uncomputed)
            Console.WriteLine("Cell A1 Formula: " + formulaCell.Formula);
            Console.WriteLine("Cell A1 Value (before calculation): " + (formulaCell.Value ?? "null"));

            // If you later need to calculate formulas manually, call:
            // workbook.CalculateFormula();

            // Keep the workbook unchanged and optionally save it
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
