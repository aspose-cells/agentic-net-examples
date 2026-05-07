using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSxcDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source SXC file (can be an existing file or a new empty workbook saved as SXC)
            string sourcePath = "Template.sxc";

            // Ensure the source file exists; if not, create a simple workbook and save it as SXC
            if (!File.Exists(sourcePath))
            {
                Workbook empty = new Workbook();
                empty.Worksheets[0].Name = "Sheet1";
                empty.Save(sourcePath, SaveFormat.Sxc);
            }

            // Load the SXC workbook
            Workbook workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // 1. Add or modify VBA macro logic
            // -------------------------------------------------
            // Check whether the workbook already contains a VBA project
            if (!workbook.HasMacro)
            {
                // To create a VBA project we need to save the workbook as a macro‑enabled format first,
                // then reload it. This is a known requirement of Aspose.Cells.
                string tempMacroPath = "TempMacro.xlsm";
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);
            }

            // Add a new procedural module named "DemoModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "DemoModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // Set the VBA code for the module
            module.Codes =
                "Sub ShowMessage()\n" +
                "    MsgBox \"Hello from Aspose.Cells VBA!\"\n" +
                "End Sub\n" +
                "\n" +
                "Sub AddNumbers()\n" +
                "    Dim a As Integer, b As Integer\n" +
                "    a = 5\n" +
                "    b = 10\n" +
                "    MsgBox \"Sum = \" & (a + b)\n" +
                "End Sub";

            // -------------------------------------------------
            // 2. Optionally remove macros (demonstration)
            // -------------------------------------------------
            // Uncomment the following line to strip all VBA from the workbook
            // workbook.RemoveMacro();

            // -------------------------------------------------
            // 3. Save the workbook back to SXC format
            // -------------------------------------------------
            string outputPath = "Result.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            Console.WriteLine($"Workbook processed. Output saved to: {outputPath}");
        }
    }
}