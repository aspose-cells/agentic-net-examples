// Title: Remove a VBA module from an Excel workbook with Aspose.Cells for .NET
// Description: Creates a workbook, adds a VBA module to the first worksheet, optionally writes a simple macro, deletes the module using Modules.RemoveAt by index, and saves the file as a macro‑enabled .xlsm to persist the change.
// Keywords: Aspose.Cells VBA module removal | Modules.RemoveAt C# | delete VBA module Aspose | save macro‑enabled workbook .xlsm | Aspose.Cells VbaProject example
// Common Searches: Aspose.Cells remove VBA module by index | How to delete a VBA module in C# using Aspose.Cells | Modules.RemoveAt usage Aspose.Cells .NET | Save workbook after VBA module removal | C# code to strip VBA modules from .xlsm
// Developer Intent: Programmatically delete a specific VBA module from a workbook and ensure the modification is written to a macro‑enabled file.
// Use Cases: Clean up temporary or unwanted VBA modules before distributing a macro‑enabled workbook. | Enforce security policies by stripping generated macros after automated processing. | Maintain only approved VBA code in archived .xlsm files.
// AI Prompts: Write C# code that lists all VBA modules in a workbook, removes one by its name, and saves the result as .xlsm using Aspose.Cells. | Explain the error handling and index rules for Modules.RemoveAt in Aspose.Cells. | Show how to confirm that a VBA module has been successfully removed after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Creates a workbook, adds a VBA module to the first worksheet, optionally writes a simple macro, deletes the module using Modules.RemoveAt by index, and saves the file as a macro‑enabled .xlsm to persist the change.
class RemoveVbaModuleDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a VBA module associated with the worksheet
        int moduleIndex = workbook.VbaProject.Modules.Add(worksheet);

        // (Optional) Add some VBA code to the module
        workbook.VbaProject.Modules[moduleIndex].Codes = 
            "Sub Test()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        // Remove the module using RemoveAt (by index)
        workbook.VbaProject.Modules.RemoveAt(moduleIndex);

        // Save the workbook as a macro‑enabled file to persist changes
        workbook.Save("WorkbookAfterModuleRemoval.xlsm", SaveFormat.Xlsm);
    }
}
