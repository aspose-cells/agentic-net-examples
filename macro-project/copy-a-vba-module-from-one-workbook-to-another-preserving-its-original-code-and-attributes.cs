// Title: Copy VBA Module Between .xlsm Workbooks Using Aspose.Cells for C#
// Description: Demonstrates how to load a macro‑enabled workbook, retrieve a specific VbaModule from its VbaProject, create an equivalent module in a new workbook, copy the module's code and attributes, and save the result as an .xlsm file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy VBA module | C# copy VbaModule | transfer macro code between workbooks | VbaProject duplicate module | macro‑enabled workbook Aspose.Cells | programmatic VBA module copy | Aspose.Cells .NET VBA automation
// Common Searches: how to copy a VBA module with Aspose.Cells C# | Aspose.Cells duplicate VbaModule between .xlsm files | C# copy macro module programmatically | preserve VBA code attributes when copying workbooks | Aspose.Cells VbaProject example for module transfer
// Developer Intent: Programmatically copy a selected VBA module from a source .xlsm workbook to a new macro‑enabled workbook while retaining its code and metadata.
// Use Cases: Apply a standard set of macros from a template to generated reports. | Deploy custom automation macros across multiple workbooks in a batch process. | Create a clean workbook that inherits existing VBA functionality for downstream users.
// AI Prompts: Generate C# code that copies all VBA modules from one workbook to another using Aspose.Cells, with error handling for missing modules. | Show how to copy a VBA module and then update its internal references (e.g., workbook names) after the transfer. | Explain how Aspose.Cells preserves module attributes such as description, references, and protection when copying a VbaModule.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to load a macro‑enabled workbook, retrieve a specific VbaModule from its VbaProject, create an equivalent module in a new workbook, copy the module's code and attributes, and save the result as an .xlsm file with Aspose.Cells for .NET.
class CopyVbaModuleDemo
{
    static void Main()
    {
        try
        {
            // Path to the source workbook that contains the VBA module.
            string sourcePath = "Source.xlsm";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook.
            Workbook destinationWorkbook = new Workbook();

            // Access the VBA projects of both workbooks.
            VbaProject sourceProject = sourceWorkbook.VbaProject;
            VbaProject destinationProject = destinationWorkbook.VbaProject;

            // Identify the module to copy (by name). Adjust the name as needed.
            string moduleNameToCopy = "MyModule";

            // Ensure the source module exists.
            VbaModule sourceModule = sourceProject.Modules[moduleNameToCopy];
            if (sourceModule == null)
            {
                Console.WriteLine($"Module '{moduleNameToCopy}' not found in source workbook.");
                return;
            }

            // Add a new module to the destination project with the same type and name.
            int newModuleIndex = destinationProject.Modules.Add(sourceModule.Type, sourceModule.Name);
            VbaModule destinationModule = destinationProject.Modules[newModuleIndex];

            // Copy the VBA code from the source module to the new module.
            destinationModule.Codes = sourceModule.Codes;

            // Save the destination workbook as a macro‑enabled file.
            string destinationPath = "Destination.xlsm";
            destinationWorkbook.Save(destinationPath, SaveFormat.Xlsm);
            Console.WriteLine($"VBA module copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
