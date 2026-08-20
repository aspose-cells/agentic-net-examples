// Title: Load an Excel workbook with Aspose.Cells and strip VBA macros using a custom LoadFilter (C#)
// Description: Shows how to configure LoadOptions with a custom LoadFilter that disables the VBA flag, load a macro‑enabled .xlsm file, confirm workbook.HasMacro is false, and optionally save the workbook as a macro‑free .xlsx.
// Keywords: Aspose.Cells LoadOptions | exclude VBA macros | custom LoadFilter | Workbook.HasMacro false | macro‑free Excel file | C# Aspose.Cells example | strip VBA from workbook | load .xlsm without macros | convert .xlsm to .xlsx | security remove Excel macros
// Common Searches: Aspose.Cells load .xlsm without VBA | How to ignore VBA project when loading Excel with Aspose.Cells | Check workbook.HasMacro after loading with LoadOptions | Save macro‑free workbook using Aspose.Cells C# | Custom LoadFilter to exclude VBA in Aspose.Cells
// Developer Intent: Load a workbook while omitting its VBA project so that no macros are retained.
// Use Cases: Strip all VBA code from a macro‑enabled file before processing to meet security policies. | Validate that a loaded workbook contains no macros by checking the HasMacro property. | Convert a .xlsm workbook to a macro‑free .xlsx while preserving worksheets, formulas, and formatting.
// AI Prompts: Generate C# code that loads an .xlsm file with Aspose.Cells, excludes VBA macros via a custom LoadFilter, and verifies workbook.HasMacro is false. | Explain step‑by‑step how to implement a LoadFilter that disables the VBA data flag when loading a workbook with Aspose.Cells. | Provide a sample that converts a macro‑enabled Excel file to a macro‑free .xlsx using LoadOptions and confirms no macros remain.

using System;
using Aspose.Cells;

// Shows how to configure LoadOptions with a custom LoadFilter that disables the VBA flag, load a macro‑enabled .xlsm file, confirm workbook.HasMacro is false, and optionally save the workbook as a macro‑free .xlsx.
class Program
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter that excludes VBA projects
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new ExcludeVbaLoadFilter();

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("sample_with_macro.xlsm", loadOptions);

        // Verify that macros have been omitted
        Console.WriteLine("HasMacro after loading with exclusion: " + workbook.HasMacro);

        // Save the workbook to a macro‑free file (optional verification)
        workbook.Save("sample_without_macro.xlsx", SaveFormat.Xlsx);
    }

    // Custom LoadFilter implementation that loads all data except VBA
    private class ExcludeVbaLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load everything (All) but remove the VBA flag
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.VBA;
        }
    }
}
