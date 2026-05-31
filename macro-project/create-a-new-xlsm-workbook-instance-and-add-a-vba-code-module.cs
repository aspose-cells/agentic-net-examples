using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new procedural VBA module named "MyModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");

        // Retrieve the added module and assign VBA code to it
        VbaModule module = vbaProject.Modules[moduleIndex];
        module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Set the encoding for the VBA project (optional but recommended)
        vbaProject.Encoding = Encoding.UTF8;

        // Save the workbook as a macro‑enabled file (XLSM)
        workbook.Save("MyWorkbook.xlsm", SaveFormat.Xlsm);
    }
}