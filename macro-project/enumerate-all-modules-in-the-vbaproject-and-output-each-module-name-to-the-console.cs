// Title: C# – List all VBA module names in a macro‑enabled workbook using Aspose.Cells
// Description: Loads an .xlsm file, accesses its VbaProject, iterates the VbaModuleCollection, and writes each module's index and name to the console with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA modules | C# enumerate VbaProject modules | list VBA module names .xlsm | VbaModuleCollection iteration | read macro‑enabled workbook Aspose
// Common Searches: how to get VBA module names from xlsm using Aspose.Cells | C# code to list VBA modules in a workbook | Aspose.Cells example for enumerating VbaProject modules | retrieve VBA module collection .NET Aspose
// Developer Intent: Extract and display the names of every VBA module contained in a macro‑enabled Excel workbook.
// Use Cases: Generate an inventory of VBA modules for documentation or audit trails. | Verify required macro modules exist before further processing. | Log module names during automated testing to aid macro debugging.
// AI Prompts: Create code that writes each VBA module name to a text file instead of the console. | Show how to include the module type (standard, class, or form) alongside the name in the output. | Provide an example that filters the module list to only display names starting with "Mod_".

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an .xlsm file, accesses its VbaProject, iterates the VbaModuleCollection, and writes each module's index and name to the console with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load a macro-enabled workbook (replace with your actual file path)
        string inputPath = "input.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project within the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve the collection of VBA modules
        VbaModuleCollection modules = vbaProject.Modules;

        // Enumerate all modules and output their names to the console
        for (int i = 0; i < modules.Count; i++)
        {
            VbaModule module = modules[i];
            Console.WriteLine($"Module {i}: {module.Name}");
        }
    }
}
