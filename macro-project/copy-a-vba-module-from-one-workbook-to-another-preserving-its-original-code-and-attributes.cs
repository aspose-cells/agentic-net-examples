// Title: Copy a specific VBA module from one macro‑enabled Excel workbook to another using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to duplicate a VBA module named "Module1" from a source .xlsm file into a destination .xlsm workbook, preserving all code and attributes. | Insert a new VBA module into an existing macro‑enabled workbook and assign it the code from a module in another workbook using the Aspose.Cells VbaProject API. | Implement robust error handling in C# for missing source/destination files or absent VBA projects when transferring a VBA module between two Excel files with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy VBA module from one .xlsm to another | How to transfer a VBA module between macro-enabled Excel files using .NET | C# example for duplicating a VBA module with Aspose.Cells VbaProject | Copy VBA code from source workbook to destination workbook Aspose.Cells
// Tags: Aspose.Cells VbaModule duplication C# | macro-enabled workbook VbaProject manipulation .NET | transfer VBA code between .xlsm files using Aspose.Cells | C# Aspose.Cells insert VbaModule into workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The program loads a source and a destination macro‑enabled Excel workbook, verifies that both contain VBA projects, locates a module named "Module1" in the source, adds a new module with the same type and name to the destination, copies the source module's code into the new module, and saves the updated workbook as a new .xlsm file, handling missing files and VBA project errors.
class Program
{
    static void Main()
    {
        try
        {
            string srcPath = "source.xlsm";
            string destPath = "destination.xlsm"; // Ensure macro‑enabled format
            string outputPath = "destination_with_module.xlsm";

            // Verify that the required files exist
            if (!File.Exists(srcPath))
            {
                Console.WriteLine($"Source file \"{srcPath}\" not found.");
                return;
            }

            if (!File.Exists(destPath))
            {
                Console.WriteLine($"Destination file \"{destPath}\" not found.");
                return;
            }

            // Load the workbooks
            Workbook srcWorkbook = new Workbook(srcPath);
            Workbook destWorkbook = new Workbook(destPath);

            // Ensure source workbook contains a VBA project
            if (srcWorkbook.VbaProject == null)
            {
                Console.WriteLine("Source workbook does not contain a VBA project.");
                return;
            }

            // Ensure destination workbook contains a VBA project
            if (destWorkbook.VbaProject == null)
            {
                Console.WriteLine("Destination workbook does not contain a VBA project. Cannot add module.");
                return;
            }

            // Name of the VBA module to copy
            string moduleName = "Module1";

            // Retrieve the source module
            VbaModule srcModule = null;
            foreach (VbaModule mod in srcWorkbook.VbaProject.Modules)
            {
                if (string.Equals(mod.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    srcModule = mod;
                    break;
                }
            }

            if (srcModule == null)
            {
                Console.WriteLine($"Module \"{moduleName}\" not found in the source workbook.");
                return;
            }

            // Add a new module to the destination workbook (Add returns the index)
            int newIndex = destWorkbook.VbaProject.Modules.Add(srcModule.Type, srcModule.Name);
            VbaModule destModule = destWorkbook.VbaProject.Modules[newIndex];

            // Copy the VBA code
            destModule.Codes = srcModule.Codes;

            // Save the destination workbook as a macro‑enabled file
            destWorkbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Module \"{moduleName}\" copied successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
