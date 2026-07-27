using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCheck
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a macro‑enabled workbook in memory
            Workbook sourceWorkbook = new Workbook();
            // Add a VBA module so the workbook contains a macro
            int moduleIndex = sourceWorkbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
            VbaModule module = sourceWorkbook.VbaProject.Modules[moduleIndex];
            module.Codes = "Sub Hello()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Save the workbook to a memory stream in macro‑enabled format (XLSM)
            using (MemoryStream macroStream = new MemoryStream())
            {
                sourceWorkbook.Save(macroStream, SaveFormat.Xlsm);
                // Reset the stream position for reading
                macroStream.Position = 0;

                // Step 2: Load the workbook from the memory stream
                Workbook loadedWorkbook = new Workbook(macroStream);

                // Step 3: Verify that the workbook contains macros and at least one VBA module
                bool hasMacro = loadedWorkbook.HasMacro;
                bool hasModules = loadedWorkbook.VbaProject != null && loadedWorkbook.VbaProject.Modules.Count > 0;

                Console.WriteLine($"Workbook loaded from stream. HasMacro: {hasMacro}");
                Console.WriteLine($"Number of VBA modules: {(hasModules ? loadedWorkbook.VbaProject.Modules.Count.ToString() : "0")}");
                Console.WriteLine(hasMacro && hasModules
                    ? "Verification succeeded: workbook contains at least one VBA module."
                    : "Verification failed: workbook does not contain VBA modules.");
            }

            // Clean up
            sourceWorkbook.Dispose();
        }
    }
}