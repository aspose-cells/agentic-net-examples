using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaManagement
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTX template (macro‑free)
            string templatePath = "Template.xltx";

            Workbook workbook;

            // Load existing template or create a new workbook if the file is missing
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Ensure the workbook has a VBA project.
            if (workbook.VbaProject == null)
            {
                string tempMacroPath = "TempMacroWorkbook.xlsm";

                // Save as a macro‑enabled workbook to create a VBA project
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);

                // Reload the workbook so that the VBA project becomes available
                workbook = new Workbook(tempMacroPath);

                // Clean up the temporary file
                File.Delete(tempMacroPath);
            }

            // Add a new standard module to the VBA project
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacroModule");

            // Retrieve the added module and set its VBA code
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
            vbaModule.Codes =
                "Sub ShowMessage()\r\n" +
                "    MsgBox \"Hello from Aspose.Cells VBA!\"\r\n" +
                "End Sub";

            // Save the workbook as a macro‑enabled file (XLSM)
            string outputPath = "OutputWithMacro.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA module at: {outputPath}");
        }
    }
}