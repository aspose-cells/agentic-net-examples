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

        // Add a new procedural VBA module named "Automation"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

        // Retrieve the added module and optionally set its VBA code
        VbaModule automationModule = vbaProject.Modules[moduleIndex];
        automationModule.Codes = "Sub RunAutomation()\n    MsgBox \"Automation module loaded\"\nEnd Sub";

        // Save the workbook as a macro‑enabled file
        workbook.Save("AutomationModule.xlsm", SaveFormat.Xlsm);
    }
}