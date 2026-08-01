// Title: Create a macro‑enabled XLSM workbook and add a VBA class module in C# with Aspose.Cells
// Description: Shows how to instantiate a Workbook, access its VbaProject, add a VBA class module named MyModule, assign a simple Sub procedure, and save the file as an XLSM macro‑enabled workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | VBA class module | VbaProject | Add VBA module | macro-enabled workbook | XLSM | SaveFormat.Xlsm | programmatic Excel macros | Excel automation
// Common Searches: add VBA class module Aspose.Cells C# | create XLSM file with VBA using Aspose.Cells | save workbook as macro enabled with Aspose.Cells .NET | set VBA code in VbaModule Aspose.Cells | generate Excel file with macros programmatically C#
// Developer Intent: Programmatically generate a macro‑enabled Excel workbook and embed a VBA class module containing custom code.
// Use Cases: Distribute a template workbook that already includes predefined VBA macros for end‑users. | Automate the creation of reports that require custom VBA classes for data processing or charting. | Build a SaaS solution that produces Excel files with embedded macros tailored to each client. | Integrate VBA‑based functionality into Excel files generated from a .NET backend.
// AI Prompts: How can I add a standard VBA module instead of a class module with Aspose.Cells? | Show code to add multiple VBA modules, each with its own procedure, using Aspose.Cells for .NET. | Explain how to reference the newly added VBA class module from Excel after the XLSM file is saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Shows how to instantiate a Workbook, access its VbaProject, add a VBA class module named MyModule, assign a simple Sub procedure, and save the file as an XLSM macro‑enabled workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new class module named "MyModule" to the VBA project
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyModule");

        // Retrieve the added module by its index
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Set VBA code for the module
        module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Save the workbook as a macro‑enabled file (XLSM)
        workbook.Save("MyWorkbook.xlsm", SaveFormat.Xlsm);
    }
}
