using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the VBA project (read‑only property, but we can add modules)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a procedural (standard) VBA module named "Automation"
            // (rule: VbaModuleCollection.Add(VbaModuleType, string))
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

            // Retrieve the added module
            VbaModule automationModule = vbaProject.Modules[moduleIndex];

            // Insert a multi‑line VBA subroutine that runs when the workbook is opened.
            // Using the Auto_Open macro which Excel executes on opening a macro‑enabled file.
            automationModule.Codes =
                "Sub Auto_Open()\r\n" +
                "    ' Log workbook opening event\r\n" +
                "    MsgBox \"Workbook opened at \" & Now\r\n" +
                "End Sub";

            // Save the workbook as a macro‑enabled file (lifecycle rule: save)
            workbook.Save("AutomationWorkbook.xlsm", SaveFormat.Xlsm);
        }
    }
}