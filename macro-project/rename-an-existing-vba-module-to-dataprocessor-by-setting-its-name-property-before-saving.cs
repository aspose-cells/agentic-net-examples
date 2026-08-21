// Title: C# Example: Change VBA Module Name to DataProcessor Using Aspose.Cells
// Description: Demonstrates creating a macro‑enabled workbook, adding a VBA class module, inserting sample code, updating the module's Name property to "DataProcessor", and saving the result as an .xlsm file—all with Aspose.Cells for .NET without launching Excel.
// Keywords: Aspose.Cells | VbaModule | rename VBA module | C# Excel automation | macro‑enabled workbook | xlsm file | set Name property | programmatic VBA rename | Excel VBA project manipulation | GitHub Aspose.Cells examples
// Common Searches: Aspose.Cells set VBA module name C# | How to rename a VBA class module programmatically | Change module identifier before saving .xlsm | Update VBA module names with Aspose.Cells for .NET | Rename macro module using C# code
// Developer Intent: Assign a new identifier to an existing VbaModule (e.g., DataProcessor) before persisting the workbook.
// Use Cases: Enforce a naming convention for dynamically generated VBA modules. | Prepare workbooks for distribution with consistent macro module names. | Migrate legacy macro projects by updating outdated module identifiers.
// AI Prompts: Generate C# code that uses Aspose.Cells to rename a VBA module to a user‑specified name and save the workbook as .xlsm. | Explain the limitations, if any, of the VbaModule.Name property when working with macro‑enabled files. | Show how to loop through all VBA modules in a workbook and apply a custom naming pattern using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates creating a macro‑enabled workbook, adding a VBA class module, inserting sample code, updating the module's Name property to "DataProcessor", and saving the result as an .xlsm file—all with Aspose.Cells for .NET without launching Excel.
class RenameVbaModule
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a VBA module with an initial name
        int moduleIndex = wb.VbaProject.Modules.Add(VbaModuleType.Class, "OldName");
        VbaModule module = wb.VbaProject.Modules[moduleIndex];

        // Optional: add some VBA code to the module
        module.Codes = "Sub Test()\r\n    MsgBox \"Hello\"\r\nEnd Sub";

        // Rename the module to "DataProcessor"
        module.Name = "DataProcessor";

        // Save the workbook as a macro-enabled file
        wb.Save("RenamedModule.xlsm", SaveFormat.Xlsm);
    }
}
