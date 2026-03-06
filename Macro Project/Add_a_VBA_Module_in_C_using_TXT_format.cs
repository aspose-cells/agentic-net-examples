using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    public class AddVbaModuleExample
    {
        public static void Run()
        {
            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Access the VBA project of the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new class module named "MyMacroModule"
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyMacroModule");

            // Retrieve the added module
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];

            // Define VBA code to be placed in the module
            string vbaCode =
                "Sub ShowMessage()\r\n" +
                "    MsgBox \"Hello from Aspose.Cells VBA module!\"\r\n" +
                "End Sub";

            // Assign the VBA code to the module
            vbaModule.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file (XLSM)
            workbook.Save("VbaModuleAdded.xlsm", SaveFormat.Xlsm);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddVbaModuleExample.Run();
        }
    }
}