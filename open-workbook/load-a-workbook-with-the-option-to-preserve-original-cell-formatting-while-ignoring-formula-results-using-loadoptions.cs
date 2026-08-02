// Title: C# – Load an Excel workbook with Aspose.Cells, preserve formatting and disable formula evaluation (LoadOptions)
// Description: Demonstrates how to create a LoadOptions object, set ParsingFormulaOnOpen to false, open an XLSX file while keeping all cell styles, colors, and fonts intact, and save the workbook without calculating any formulas.
// Keywords: Aspose.Cells LoadOptions | ParsingFormulaOnOpen false | load Excel without formula calculation | preserve cell formatting Aspose | C# Excel workbook loading example | skip formula evaluation Aspose.Cells
// Common Searches: Aspose.Cells open workbook without evaluating formulas | keep original Excel styles when loading with Aspose | LoadOptions to ignore formulas in C# | preserve formatting while disabling formula parsing Aspose.Cells
// Developer Intent: Open an Excel file, retain its visual formatting, and prevent any formula calculations during the load process.
// Use Cases: Render a template workbook for reporting where only layout matters, not formula results. | Safely ingest user‑uploaded spreadsheets to extract raw data without triggering volatile or external formulas. | Copy styling from a source workbook to a new file while leaving formula outcomes untouched for later processing.
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, disables formula evaluation using LoadOptions, and saves it preserving all styles. | Explain how the ParsingFormulaOnOpen property impacts performance and memory consumption when loading large workbooks. | Suggest alternative approaches in Aspose.Cells to read cell values without evaluating formulas.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadOptionsDemo
{
    // Demonstrates how to create a LoadOptions object, set ParsingFormulaOnOpen to false, open an XLSX file while keeping all cell styles, colors, and fonts intact, and save the workbook without calculating any formulas.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Path for the resulting workbook
            string outputPath = "output.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Preserve original cell formatting (default behavior) and ignore formula parsing
            // Setting ParsingFormulaOnOpen to false skips formula evaluation during load
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point, cell values for formulas are not calculated,
            // but formatting (styles, colors, fonts, etc.) is retained.

            // Save the workbook to verify the result
            workbook.Save(outputPath);
        }
    }
}
