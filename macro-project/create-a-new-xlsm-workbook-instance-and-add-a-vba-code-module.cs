// Title: Create an XLSM workbook and insert a VBA class module with Aspose.Cells for .NET
// Description: Shows how to create a new Workbook, access its VbaProject, add a VBA class module named MyModule, embed a simple Sub routine, and save the result as a macro‑enabled XLSM file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | VBA class module | XLSM creation | C# Excel automation | VbaProject | macro enabled workbook | programmatic VBA insertion | SaveFormat.Xlsm | .NET Excel library
// Common Searches: Aspose.Cells add VBA class module C# | Create XLSM file with VBA code using Aspose.Cells | How to save macro enabled workbook .NET | Insert VBA code into Excel programmatically | Aspose.Cells VbaProject example
// Developer Intent: Add a VBA class module with code to a newly created macro‑enabled workbook and persist it as an XLSM file.
// Use Cases: Automated report generation that includes custom macros for end‑user tasks. | Building Excel templates pre‑loaded with VBA utilities for distribution across an organization. | Server‑side services that produce macro‑enabled files without requiring Microsoft Excel.
// AI Prompts: Show how to add a standard VBA module instead of a class module with Aspose.Cells. | Provide code to create multiple VBA modules, each containing a different procedure. | Explain how to reference external VBA libraries when adding modules via Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    // Shows how to create a new Workbook, access its VbaProject, add a VBA class module named MyModule, embed a simple Sub routine, and save the result as a macro‑enabled XLSM file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new VBA class module named "MyModule"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyModule");

            // Retrieve the added module and set its VBA code
            VbaModule module = vbaProject.Modules[moduleIndex];
            module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Save the workbook as a macro‑enabled file (XLSM)
            workbook.Save("MyWorkbookWithVba.xlsm", SaveFormat.Xlsm);
        }
    }
}
