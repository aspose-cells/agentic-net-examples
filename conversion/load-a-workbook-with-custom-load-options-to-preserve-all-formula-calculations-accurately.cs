// Title: C# – Load an Excel workbook with custom LoadOptions to preserve formula padding, retain unparsed data, and recalculate formulas using Aspose.Cells
// Description: Demonstrates how to configure LoadOptions (ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData), open an XLSX file, force a full formula calculation, and save the result with Aspose.Cells for .NET.
// Keywords: Aspose.Cells LoadOptions | ParsingFormulaOnOpen | PreservePaddingSpacesInFormula | KeepUnparsedData | recalculate formulas C# | load Excel with custom options | formula padding preservation
// Common Searches: Aspose.Cells load workbook with custom options | preserve spaces in Excel formulas Aspose | enable formula parsing on open .NET | keep unparsed data when loading Excel file | force formula calculation after opening workbook
// Developer Intent: Open an Excel file with specific LoadOptions to keep exact formula text and run a full calculation before saving.
// Use Cases: Maintain exact spacing and line breaks in complex formulas during conversion or analysis. | Access raw, unparsed cell data for custom processing without losing original content. | Generate up‑to‑date calculation results after loading a workbook for reporting or further manipulation.
// AI Prompts: Generate C# code that opens an XLSX file with Aspose.Cells, preserving formula padding and unparsed data, then calculates all formulas and saves the file. | Explain the effect of ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, and KeepUnparsedData on workbook loading and formula evaluation in Aspose.Cells. | Provide a step‑by‑step tutorial for using LoadOptions to retain formula formatting, trigger full calculation, and export the workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to configure LoadOptions (ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData), open an XLSX file, force a full formula calculation, and save the result with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create custom load options
        LoadOptions loadOptions = new LoadOptions();

        // Parse all formulas when the workbook is opened
        loadOptions.ParsingFormulaOnOpen = true;

        // Preserve any padding spaces or line breaks inside formulas
        loadOptions.PreservePaddingSpacesInFormula = true;

        // Keep unparsed data in memory (default is true, set explicitly for clarity)
        loadOptions.KeepUnparsedData = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure all formulas are calculated after loading
        workbook.CalculateFormula();

        // Save the workbook (can overwrite or save to a new file)
        workbook.Save("output.xlsx");
    }
}
