// Title: How to clear a VBA module and import new macro code from a .bas file into an XLSM workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace VBA module code with external .bas file | clear existing macro code in a VBA module and add new code using Aspose.Cells | programmatically update an .xlsm workbook's VBA module in .NET | load macro-enabled Excel file and modify VBA project with Aspose.Cells | read .bas file and assign to VBA module in an XLSM workbook using C#
// Tags: replace VBA module Aspose.Cells | import .bas macro into XLSM Aspose.Cells | clear VBA module C# | update macro-enabled workbook Aspose.Cells | load macro-enabled workbook C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads a macro‑enabled XLSM workbook with Aspose.Cells, verifies the VBA project, retrieves a specified module, clears its existing code, reads new macro code from an external .bas file, assigns the code to the module, and saves the workbook back as an XLSM file, handling missing files and runtime errors.
class MacroUpdater
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains the VBA project
            string workbookPath = "input.xlsm";

            // Name of the VBA module to be cleared and updated
            string moduleName = "Module1";

            // Path to the external file that holds the updated macro code
            string macroSourcePath = "updatedMacro.bas";

            // Path where the updated workbook will be saved
            string outputPath = "output.xlsm";

            // Verify required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(macroSourcePath))
                throw new FileNotFoundException($"Macro source file not found: {macroSourcePath}");

            // Load the existing workbook preserving macros (XLSX format works for .xlsm files)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject == null)
                throw new InvalidOperationException("The workbook does not contain a VBA project.");

            // Access the VBA project within the workbook
            var vbaProject = workbook.VbaProject;

            // Retrieve the target module by its name
            var targetModule = vbaProject.Modules[moduleName];
            if (targetModule == null)
                throw new InvalidOperationException($"VBA module '{moduleName}' not found in the workbook.");

            // Clear any existing code in the module
            targetModule.Codes = string.Empty;

            // Read the new macro code from the external source file
            string newMacroCode = File.ReadAllText(macroSourcePath);

            // Insert the new macro code into the module
            targetModule.Codes = newMacroCode;

            // Save the workbook, ensuring the macro-enabled format is used
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
