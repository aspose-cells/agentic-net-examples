using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural VBA module named "Automation"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

        // Retrieve the newly added module
        VbaModule automationModule = vbaProject.Modules[moduleIndex];

        // (Optional) Add some VBA code to the module
        automationModule.Codes = "Sub AutomationMacro()\n    MsgBox \"Automation module loaded\"\nEnd Sub";

        // Save the workbook as a macro‑enabled file
        workbook.Save("AutomationModule.xlsm", SaveFormat.Xlsm);
    }
}