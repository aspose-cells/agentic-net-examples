using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class RemoveVbaModuleByIndexDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new macro‑enabled workbook
            Workbook workbook = new Workbook();

            // Access the VBA project of the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new procedural VBA module with a specific name
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "SampleModule");

            // Add some VBA code to the module
            vbaProject.Modules[moduleIndex].Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello World\"\r\nEnd Sub";

            // Remove the module by its index
            vbaProject.Modules.RemoveAt(moduleIndex);

            // Save the workbook as a macro‑enabled file
            workbook.Save("WorkbookAfterModuleRemoval.xlsm", SaveFormat.Xlsm);
        }
    }
}