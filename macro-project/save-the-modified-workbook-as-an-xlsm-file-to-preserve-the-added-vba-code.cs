// Title: Save Workbook with VBA Class Module as Macro‑Enabled XLSM using Aspose.Cells for .NET
// Description: Shows how to create a new Workbook, add a VBA class module called DemoModule containing a HelloWorld macro via Aspose.Cells VbaProject, and save the file as XLSM so the VBA code is preserved.
// Keywords: Aspose.Cells | C# | .NET | VBA | macro‑enabled workbook | XLSM | VbaProject | add VBA module | save workbook with VBA | programmatic Excel macro
// Common Searches: add VBA class module with Aspose.Cells | save Excel file as macro enabled XLSM .NET | Aspose.Cells VbaProject example | embed VBA code in workbook using C# | how to preserve VBA when saving with Aspose.Cells
// Developer Intent: Embed a VBA class module into a new workbook and persist it by saving as a macro‑enabled XLSM file.
// Use Cases: Create template workbooks that ship with predefined macros for end‑users. | Inject custom VBA functions into financial or reporting spreadsheets before distribution. | Automate generation of macro‑enabled workbooks as part of a CI/CD pipeline for Excel add‑ins.
// AI Prompts: Generate C# code that adds several VBA modules (standard, class, and form) and saves the workbook as .xlsm with Aspose.Cells. | Explain how to load an existing workbook, modify its VBA code, and re‑save it as a macro‑enabled file using Aspose.Cells. | Show how to set a password on the VBA project and preserve it when saving a macro‑enabled workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    // Shows how to create a new Workbook, add a VBA class module called DemoModule containing a HelloWorld macro via Aspose.Cells VbaProject, and save the file as XLSM so the VBA code is preserved.
    public class SaveWorkbookWithVba
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the VBA project (creates a project when saved as macro‑enabled)
                VbaProject vbaProject = workbook.VbaProject;

                // Add a new VBA module of type Class named "DemoModule"
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");

                // Retrieve the added module and set its VBA code
                VbaModule module = vbaProject.Modules[moduleIndex];
                module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

                // Save the workbook as a macro‑enabled file to preserve the VBA code
                workbook.Save("WorkbookWithVba.xlsm", SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SaveWorkbookWithVba.Run();
        }
    }
}
