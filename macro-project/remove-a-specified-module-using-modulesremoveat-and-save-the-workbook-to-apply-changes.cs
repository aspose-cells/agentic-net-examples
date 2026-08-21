// Title: Remove a VBA module with Modules.RemoveAt in Aspose.Cells (C#) and save a macro‑free workbook
// Description: Demonstrates how to create a workbook, add a procedural VBA module, delete the module by its index using VbaProject.Modules.RemoveAt, and then save the file as a macro‑free XLSX with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# remove VBA module | Modules.RemoveAt example | delete VBA module programmatically | save workbook without macros | Aspose.Cells VbaProject | macro‑free Excel file | C# Excel automation | GitHub Aspose.Cells VBA example | Aspose.Cells workbook save format | remove VBA code Aspose
// Common Searches: How to use Modules.RemoveAt in Aspose.Cells C# | Remove VBA module by index and save as .xlsx | Aspose.Cells example for deleting VBA modules | Save Excel file without macros using Aspose.Cells | C# code to strip VBA from a workbook
// Developer Intent: The developer needs to delete a specific VBA module from an Excel workbook and persist the change by saving the file without any macros.
// Use Cases: Strip test or temporary VBA modules from automatically generated reports before distribution. | Create compliance‑ready workbooks by removing all macro code prior to archiving or sharing. | Clean up legacy macro projects by programmatically deleting unwanted modules.
// AI Prompts: Generate C# code that removes a VBA module by its index using Modules.RemoveAt and saves the workbook as a macro‑free .xlsx with Aspose.Cells. | Show how to list all VBA modules in a workbook, find the index of "DemoModule", delete it with RemoveAt, and keep the remaining modules intact. | Explain how to verify that a VBA module has been successfully removed after saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to create a workbook, add a procedural VBA module, delete the module by its index using VbaProject.Modules.RemoveAt, and then save the file as a macro‑free XLSX with Aspose.Cells for .NET.
class RemoveVbaModuleDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project within the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural VBA module named "DemoModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "DemoModule");

        // (Optional) Add some VBA code to the newly created module
        vbaProject.Modules[moduleIndex].Codes = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub";

        // Remove the module by its name using the Remove(string) method
        vbaProject.Modules.Remove("DemoModule");

        // Save the workbook after the removal (saved as a macro‑free file)
        workbook.Save("RemovedModule.xlsx", SaveFormat.Xlsx);
    }
}
