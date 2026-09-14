// Title: Load an Excel workbook with custom LoadOptions to preserve formula padding and auto‑parse formulas on open using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells LoadOptions to enable ParsingFormulaOnOpen and PreservePaddingSpacesInFormula, then open an .xlsx file, recalculate its formulas, and save the result. | Create a Workbook with LoadOptions that turn off KeepUnparsedData, keep formula whitespace, force formula parsing at load time, and perform a full formula calculation before saving.
// Common Searches: Aspose.Cells how to keep spaces inside Excel formulas when loading a workbook | load workbook with ParsingFormulaOnOpen option in C# | disable KeepUnparsedData for faster loading in Aspose.Cells | recalculate all formulas after opening an Excel file with Aspose.Cells | preserve formula line breaks using LoadOptions in Aspose.Cells .NET
// Tags: custom LoadOptions for workbook loading Aspose.Cells | enable ParsingFormulaOnOpen in LoadOptions | preserve padding spaces in Excel formulas | disable KeepUnparsedData for faster load | recalculate formulas after workbook load

using System;
using Aspose.Cells;

// Demonstrates configuring LoadOptions to parse formulas on opening, preserve whitespace inside formulas, skip storing unparsed data, loading an Excel workbook, recalculating all formulas, and saving the updated file.
class LoadWorkbookWithCustomOptions
{
    static void Main()
    {
        // Create custom load options
        LoadOptions loadOptions = new LoadOptions();

        // Parse all formulas when the workbook is opened
        loadOptions.ParsingFormulaOnOpen = true;

        // Preserve any padding spaces or line breaks inside formulas
        loadOptions.PreservePaddingSpacesInFormula = true;

        // Do not keep unparsed data in memory (optional, improves performance)
        loadOptions.KeepUnparsedData = false;

        // Load the workbook using the custom options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Recalculate formulas to ensure values are up‑to‑date after loading
        workbook.CalculateFormula();

        // Save the workbook (can be the same file or a new one)
        workbook.Save("output.xlsx");
    }
}
