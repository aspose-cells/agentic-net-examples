// Title: Add a Procedural VBA Module “Automation” to an XLSM Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, access its VbaProject, add a procedural VBA module called Automation, inject a simple Sub routine, and save the file as a macro‑enabled XLSM using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA module | add procedural VBA module C# | VbaProject Aspose.Cells | macro-enabled workbook .NET | save as Xlsm Aspose | C# Excel automation Aspose | Aspose.Cells VbaModule example
// Common Searches: Aspose.Cells add VBA module C# | Create macro‑enabled Excel file with Aspose | How to insert a VBA module programmatically using Aspose.Cells | Save workbook as XLSM using Aspose.Cells for .NET | Add procedural VBA code to Excel with Aspose
// Developer Intent: Insert a new procedural VBA module named Automation into a workbook and export it as an XLSM file using Aspose.Cells for .NET.
// Use Cases: Generate a template workbook that ships with a standard Automation macro for end‑users. | Batch‑process multiple workbooks to embed common VBA routines automatically. | Create macro‑enabled reporting files that contain predefined data‑validation procedures.
// AI Prompts: Write C# code with Aspose.Cells that adds a VBA module called Automation, sets a Sub showing a message box, and saves the workbook as XLSM. | Show how to read, modify, and replace the code of an existing VBA module in an Excel file using Aspose.Cells for .NET. | Explain how to list all VBA modules in a workbook and extract their source code with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to create a workbook, access its VbaProject, add a procedural VBA module called Automation, inject a simple Sub routine, and save the file as a macro‑enabled XLSM using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new procedural VBA module named "Automation"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

        // Retrieve the added module and optionally set its VBA code
        VbaModule module = vbaProject.Modules[moduleIndex];
        module.Codes = "Sub AutomationMacro()\r\n    MsgBox \"Automation module loaded\"\r\nEnd Sub";

        // Save the workbook as a macro‑enabled file
        workbook.Save("AutomationModule.xlsm", SaveFormat.Xlsm);
    }
}
