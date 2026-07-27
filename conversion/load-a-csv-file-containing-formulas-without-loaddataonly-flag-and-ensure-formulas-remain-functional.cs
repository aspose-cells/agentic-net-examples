// Title: Import a CSV containing Excel formulas with Aspose.Cells for .NET (C#) and keep them functional
// Description: Creates a Workbook, sets TxtLoadOptions.HasFormula and ParsingFormulaOnOpen, imports the CSV via ImportCSV, parses any remaining formulas, calculates results, and saves the file as XLSX—all without using the LoadDataOnly flag.
// Keywords: Aspose.Cells CSV import formulas | TxtLoadOptions HasFormula true | ParsingFormulaOnOpen | Workbook.ParseFormulas | Workbook.CalculateFormula | ImportCSV C# example | convert CSV to XLSX with formulas | .NET Excel formula parsing | load CSV without LoadDataOnly
// Common Searches: Aspose.Cells import CSV with formulas C# | Enable formula detection when loading CSV Aspose.Cells | Parse and calculate formulas after ImportCSV | How to keep formulas alive when converting CSV to Excel | TxtLoadOptions example for formula CSV
// Developer Intent: Load a CSV that contains formula strings, have Aspose.Cells recognize and evaluate those formulas, and output a fully calculated Excel workbook.
// Use Cases: Transform a financial‑report CSV that includes "=SUM(A1:B1)" cells into an .xlsx with computed totals. | Migrate data exports from a third‑party system that embed Excel formulas into a workbook while preserving the calculations. | Automate the conversion of CSV‑based templates containing "=IF(...)" logic into ready‑to‑use Excel files.
// AI Prompts: Show C# code that imports a CSV with embedded Excel formulas using Aspose.Cells, parses them, and saves the workbook as XLSX. | Explain how TxtLoadOptions.HasFormula and ParsingFormulaOnOpen work together with Workbook.ParseFormulas. | Provide a step‑by‑step guide to convert a formula‑rich CSV to Excel without setting LoadDataOnly in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvFormulaDemo
{
    // Creates a Workbook, sets TxtLoadOptions.HasFormula and ParsingFormulaOnOpen, imports the CSV via ImportCSV, parses any remaining formulas, calculates results, and saves the file as XLSX—all without using the LoadDataOnly flag.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file that contains formulas (e.g., cells with values like "=SUM(A1:B1)")
            string csvPath = "input.csv";

            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Configure TxtLoadOptions to treat text starting with "=" as formulas
            TxtLoadOptions txtOptions = new TxtLoadOptions
            {
                // Use comma as the delimiter for CSV
                SeparatorString = ",",
                // Enable formula detection
                HasFormula = true,
                // Optional: convert numeric strings to numbers
                ConvertNumericData = true,
                // Ensure formulas are parsed when the workbook is opened
                ParsingFormulaOnOpen = true
            };

            // Import the CSV data into the first worksheet starting at cell A1 (load rule)
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, txtOptions, 0, 0);

            // Parse any formulas that were not parsed during import
            workbook.ParseFormulas(false);

            // Calculate the formulas so that cell values are updated
            workbook.CalculateFormula();

            // Save the workbook to an Excel file (save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"CSV file loaded and formulas processed. Output saved to '{outputPath}'.");
        }
    }
}
