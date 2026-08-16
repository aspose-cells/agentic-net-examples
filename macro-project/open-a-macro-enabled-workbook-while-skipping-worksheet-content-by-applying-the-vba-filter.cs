// Title: Open an .xlsm workbook with Aspose.Cells for .NET using LoadFilter to load only VBA and structure
// Description: Demonstrates loading a macro‑enabled Excel file (.xlsm) in Aspose.Cells for .NET by configuring LoadOptions with a LoadFilter that includes only the VBA project and workbook structure, thereby skipping cell data, formulas, charts, and other sheet content, and confirms macro presence via Workbook.HasMacro.
// Keywords: Aspose.Cells | .xlsm | LoadFilter | VBA | skip worksheet data | load workbook structure | C# | LoadOptions | macro detection
// Common Searches: Aspose.Cells open .xlsm without loading sheet data | LoadFilter VBA only Aspose.Cells | C# load macro-enabled workbook structure only | Check if Excel file has macro using Aspose.Cells | Skip cell content when loading .xlsm with Aspose
// Developer Intent: Load an .xlsm file while excluding all worksheet content, retaining only the VBA project and workbook layout.
// Use Cases: Quickly verify the presence of macros in large batches of .xlsm files without the overhead of loading cell values. | Extract or analyze VBA modules from a workbook while keeping memory usage low. | Obtain sheet names and workbook hierarchy of a macro workbook for indexing or reporting.
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, using LoadOptions to load only VBA and the workbook structure, then list all macro modules. | Show how to save a workbook that was loaded with only VBA and structure back to a new .xlsm file while preserving the macros. | Explain how to combine LoadFilter options to load VBA, structure, and defined names without loading any cell data in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroLoadExample
{
    // Demonstrates loading a macro‑enabled Excel file (.xlsm) in Aspose.Cells for .NET by configuring LoadOptions with a LoadFilter that includes only the VBA project and workbook structure, thereby skipping cell data, formulas, charts, and other sheet content, and confirms macro presence via Workbook.HasMacro.
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook
            string macroFilePath = "sample_with_macro.xlsm";

            // Create LoadOptions and set a LoadFilter that loads only VBA projects and the workbook structure.
            // This skips loading cell data, formulas, charts, etc., while keeping the macro information.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new LoadFilter(
                LoadDataFilterOptions.VBA |        // Load VBA projects (macros)
                LoadDataFilterOptions.Structure   // Load only the workbook structure (no sheet content)
            );

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(macroFilePath, loadOptions);

            // Verify that the macro was loaded
            Console.WriteLine("Workbook has macro: " + workbook.HasMacro);
        }
    }
}
