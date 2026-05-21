using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook() constructor)
            Workbook workbook = new Workbook();

            // Access the VBA project (read‑only property, but we can add modules after saving as macro‑enabled)
            // Ensure the workbook is macro‑enabled by saving it as Xlsm and reloading
            string tempPath = "temp.xlsm";
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            System.IO.File.Delete(tempPath);

            // Add a new VBA module to the project
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "TestModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Codes = "Sub Test()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Save the workbook as a macro‑enabled file to preserve the VBA code
            string outputPath = "VbaEnabledWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA code at: {outputPath}");
        }
    }
}