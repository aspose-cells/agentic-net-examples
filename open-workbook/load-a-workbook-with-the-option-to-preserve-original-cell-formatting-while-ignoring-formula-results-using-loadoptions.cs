// Title: Load an Excel workbook with Aspose.Cells while preserving formatting and disabling formula evaluation (LoadOptions)
// Description: Shows C# code that creates a LoadOptions object, sets ParsingFormulaOnOpen = false, loads an XLSX file without calculating formulas, keeps the original cell styles, displays the raw formula/value of A1, and saves the workbook.
// Keywords: Aspose.Cells LoadOptions | ParsingFormulaOnOpen | preserve cell formatting | disable formula calculation | load workbook without evaluating formulas | C# Excel loading example | .NET Aspose.Cells raw formulas | keep original styles Aspose | Excel template loading | skip formula parsing
// Common Searches: Aspose.Cells load workbook without calculating formulas | How to keep original formatting when opening Excel with Aspose | ParsingFormulaOnOpen C# example | LoadOptions preserve cell styles Aspose.Cells | Open Excel file as raw data Aspose .NET
// Developer Intent: Open an Excel file with Aspose.Cells, retain every original cell style, and prevent any formula from being evaluated during the load operation.
// Use Cases: Load a pre‑formatted template, modify only data cells, and save without altering existing formulas or styles. | Extract raw formula strings for auditing or migration while leaving the workbook unchanged. | Read a workbook in a data‑processing pipeline where calculation results are irrelevant but formatting must stay intact.
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, sets ParsingFormulaOnOpen to false, and saves it preserving all formatting. | Explain how to read raw formulas from a workbook after disabling formula parsing with LoadOptions. | Combine LoadOptions with other Aspose.Cells settings to maintain number formats, conditional formatting, and cell styles while skipping calculations.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadOptionsDemo
{
    // Shows C# code that creates a LoadOptions object, sets ParsingFormulaOnOpen = false, loads an XLSX file without calculating formulas, keeps the original cell styles, displays the raw formula/value of A1, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Do not parse formulas on load – this keeps the original formula strings
            // and prevents calculation of their results.
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook with the specified options.
            // This preserves original cell formatting and skips formula evaluation.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example: display a formula and its (unchanged) value in the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine("Cell A1 Formula: " + cell.Formula);
            Console.WriteLine("Cell A1 Value (raw): " + cell.Value);

            // Save the workbook to a new file.
            string outputPath = "output_preserve_format.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
