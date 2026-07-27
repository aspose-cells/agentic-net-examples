// Title: Load an Excel workbook without evaluating formulas using Aspose.Cells LoadOptions (.NET)
// Description: Demonstrates how to create a LoadOptions object with ParsingFormulaOnOpen set to false, open a workbook so formulas stay as raw strings, read a cell's formula and its unevaluated value, optionally trigger calculation later, and save the file while preserving the original formulas.
// Keywords: Aspose.Cells LoadOptions | disable formula calculation | ParsingFormulaOnOpen false | open workbook without evaluating formulas | read raw formula string | C# Excel loading | preserve formulas Aspose | manual formula calculation
// Common Searches: Aspose.Cells load workbook without calculating formulas | How to keep formulas unchanged when opening Excel with Aspose | Read formula text only in C# Aspose.Cells | Disable automatic formula evaluation Aspose.Cells | Load Excel file as raw formulas .NET
// Developer Intent: Open an Excel file, keep formulas intact, and prevent automatic calculation.
// Use Cases: Extract the formula text from cells without triggering any computation. | Defer formula evaluation until a specific point in the workflow. | Modify and save a workbook while guaranteeing that original formulas remain unchanged.
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, disables formula parsing on open, reads the formula from cell B2, and saves the file without altering any formulas. | Show how to load a workbook with ParsingFormulaOnOpen set to false, then later calculate all formulas programmatically using Aspose.Cells. | Provide an example that prints both the raw formula string and the default (unevaluated) cell value after loading a workbook with formulas disabled.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadFormulaDemo
{
    // Demonstrates how to create a LoadOptions object with ParsingFormulaOnOpen set to false, open a workbook so formulas stay as raw strings, read a cell's formula and its unevaluated value, optionally trigger calculation later, and save the file while preserving the original formulas.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string inputFile = "input.xlsx";

            // Create LoadOptions and disable formula parsing on open
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false; // Formulas will be loaded as raw strings, not calculated

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: read a cell that contains a formula
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine("Cell A1 Formula: " + cell.Formula);
            Console.WriteLine("Cell A1 Value (should be default, not calculated): " + (cell.Value ?? "null"));

            // If you later need to calculate formulas manually, call:
            // workbook.CalculateFormula();

            // Optionally save the workbook (formulas remain unchanged)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
