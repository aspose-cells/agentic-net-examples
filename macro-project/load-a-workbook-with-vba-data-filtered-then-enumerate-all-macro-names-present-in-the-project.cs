// Title: List VBA Module Names in an .xlsm Workbook Using Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, enables macro processing, checks for VBA presence, accesses the VbaProject, iterates the VbaModuleCollection, and prints each module's Name – effectively enumerating all macro containers in the file.
// Keywords: Aspose.Cells | C# | .xlsm | VBA modules | VbaProject | enumerate macros | list macro names | enable macros Aspose | Workbook.HasMacro
// Common Searches: list VBA modules Aspose.Cells C# | enumerate macro names in .xlsm file | how to get VbaModule names with Aspose | check workbook contains macros Aspose.Cells | access VBA project from .xlsm using .NET
// Developer Intent: Retrieve the names of all VBA modules (macros) contained in a macro‑enabled workbook.
// Use Cases: Verify required macros exist before processing the workbook. | Create an inventory of macro names for documentation or compliance audits. | Select specific modules for dynamic execution or code‑generation tasks.
// AI Prompts: Generate C# code with Aspose.Cells that extracts the full source code of each VBA module in an .xlsm file. | Show how to filter VBA modules by a naming pattern and list only the matching macro names. | Provide an example that disables macro processing after enumerating modules to reduce memory usage.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro‑enabled workbook, enables macro processing, checks for VBA presence, accesses the VbaProject, iterates the VbaModuleCollection, and prints each module's Name – effectively enumerating all macro containers in the file.
class Program
{
    static void Main()
    {
        // Load the macro-enabled workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsm");

        // Enable macros for the loaded workbook (required for VBA access)
        workbook.Settings.EnableMacros = true;

        // Check if the workbook actually contains VBA macros
        if (!workbook.HasMacro)
        {
            Console.WriteLine("The workbook does not contain any macros.");
            return;
        }

        // Get the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Get the collection of VBA modules (each module can contain one or more macros)
        VbaModuleCollection modules = vbaProject.Modules;

        Console.WriteLine($"Total VBA modules: {modules.Count}");

        // Enumerate each module and output its name (module name is typically the macro container)
        for (int i = 0; i < modules.Count; i++)
        {
            VbaModule module = modules[i];

            // The VbaModule class exposes a Name property that holds the module's identifier
            Console.WriteLine($"Module {i + 1}: {module.Name}");
        }
    }
}
