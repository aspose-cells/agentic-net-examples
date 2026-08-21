// Title: Load Excel workbook with custom LoadOptions to preserve formulas – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to configure Aspose.Cells LoadOptions (ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData, CultureInfo) to load an .xlsx, keep original formula formatting, recalculate all formulas, and save the workbook.
// Keywords: Aspose.Cells | LoadOptions | ParsingFormulaOnOpen | PreservePaddingSpacesInFormula | KeepUnparsedData | CultureInfo | formula calculation | Excel loading .NET | preserve formula formatting | locale-specific Excel | C#
// Common Searches: Aspose.Cells load workbook with formula parsing enabled | keep padding spaces in Excel formulas Aspose.Cells | LoadOptions KeepUnparsedData example | set CultureInfo when loading Excel with Aspose.Cells .NET | calculate all formulas after loading workbook Aspose.Cells
// Developer Intent: Load an Excel file using Aspose.Cells with custom LoadOptions that parse formulas on open, retain spacing/line‑breaks, keep unparsed data, apply a specific CultureInfo, then optionally recalculate and save the workbook.
// Use Cases: Open a workbook for analysis while preserving the exact formula text (spacing, line breaks). | Ensure all formulas are evaluated immediately after loading to obtain current cell values. | Maintain unparsed data for round‑trip editing without losing hidden or custom information. | Handle locale‑dependent numbers and dates by specifying the appropriate CultureInfo during load. | Modify or recalculate a workbook and save it without altering the original formula layout.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions to enable ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData, set CultureInfo to 'en-US', calculate all formulas, and save the workbook. | Explain the impact of ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData, and CultureInfo on formula handling in Aspose.Cells. | Provide a step‑by‑step tutorial for loading an Excel file, preserving formula formatting, recalculating formulas, and saving the result with Aspose.Cells for .NET.

using System;
using System.Globalization;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells LoadOptions (ParsingFormulaOnOpen, PreservePaddingSpacesInFormula, KeepUnparsedData, CultureInfo) to load an .xlsx, keep original formula formatting, recalculate all formulas, and save the workbook.
class LoadWorkbookWithCustomOptions
{
    static void Main()
    {
        // Create load options and configure them to preserve formula calculations
        LoadOptions loadOptions = new LoadOptions();

        // Parse formulas when the workbook is opened to ensure they are available for calculation
        loadOptions.ParsingFormulaOnOpen = true;

        // Preserve any padding spaces or line breaks inside formulas
        loadOptions.PreservePaddingSpacesInFormula = true;

        // Keep unparsed data in memory (default is true) – useful when the workbook will be saved later
        loadOptions.KeepUnparsedData = true;

        // Set culture info if the workbook contains locale‑specific data
        loadOptions.CultureInfo = new CultureInfo("en-US");

        // Load the workbook using the custom load options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Optionally calculate all formulas to ensure cell values are up‑to‑date
        workbook.CalculateFormula();

        // Save the workbook after loading and optional calculation
        workbook.Save("output.xlsx");
    }
}
