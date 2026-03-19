using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class InsertVbaModule
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new procedural module named "MyMacroModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacroModule");

        // Retrieve the added module
        VbaModule module = vbaProject.Modules[moduleIndex];

        // VBA code with TAB indentation
        string vbaCode = "Sub ShowMessage()\r\n" +
                         "\tMsgBox \"Hello from Aspose.Cells VBA!\"\r\n" +
                         "End Sub";

        // Set the code for the module
        module.Codes = vbaCode;

        // Save the workbook as a macro‑enabled file
        workbook.Save("InsertedVbaModule.xlsm", SaveFormat.Xlsm);
    }
}