using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ModifyVbaAndSaveXps
{
    static void Main()
    {
        // Load a macro‑enabled workbook (XLSM)
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project; it will be null if the workbook has no macros
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject != null)
        {
            // Add a new procedural module (or you could modify an existing one)
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "ModifiedModule");
            VbaModule module = vbaProject.Modules[moduleIndex];

            // Set the VBA code for the new module
            module.Codes = "Sub ModifiedMacro()\r\n    MsgBox \"Macro modified by Aspose.Cells\"\r\nEnd Sub";
        }

        // Create XPS save options and configure desired settings
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            OnePagePerSheet = true,   // each sheet on a single page
            DefaultFont = "Arial"     // fallback font for Unicode characters
        };

        // Save the workbook as XPS using the specified options
        workbook.Save("output.xps", xpsOptions);
    }
}