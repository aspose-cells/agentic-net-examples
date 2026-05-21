using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the VBA project of the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a procedural (standard) VBA module named "Automation"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "Automation");

            // Retrieve the added module
            VbaModule automationModule = vbaProject.Modules[moduleIndex];

            // Define a multi‑line VBA subroutine that logs the workbook opening event
            string vbaCode = 
                "Sub LogWorkbookOpen()\r\n" +
                "    ' This subroutine can be called from Workbook_Open event\r\n" +
                "    MsgBox \"Workbook opened at \" & Now\r\n" +
                "End Sub";

            // Assign the VBA code to the module
            automationModule.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file (Xlsm) so the VBA project is retained
            workbook.Save("AutomationModuleDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}