using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLS file (may or may not contain macros)
            string sourcePath = "SourceWorkbook.xls";

            // Load or create the workbook
            Workbook workbook;
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook(); // creates a new empty workbook
                // Save as macro‑enabled to initialise the VBA project
                string tempMacroPath = "temp.xlsm";
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);
            }

            // Ensure the workbook has a VBA project.
            if (workbook.VbaProject == null)
            {
                string tempMacroPath = "temp.xlsm";
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // -------------------------------------------------
            // Example 1: Add a new procedural module and set its code
            // -------------------------------------------------
            int procModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "UtilityModule");
            VbaModule procModule = vbaProject.Modules[procModuleIndex];
            procModule.Codes =
                "Sub ShowMessage()\r\n" +
                "    MsgBox \"Hello from Aspose.Cells VBA!\"\r\n" +
                "End Sub";

            // -------------------------------------------------
            // Example 2: Modify an existing module (if any)
            // -------------------------------------------------
            if (vbaProject.Modules.Count > 0)
            {
                VbaModule existingModule = vbaProject.Modules[0];
                string additionalCode =
                    "\r\nSub AddedMacro()\r\n" +
                    "    MsgBox \"This macro was added programmatically.\"\r\n" +
                    "End Sub";
                existingModule.Codes = existingModule.Codes + additionalCode;
            }

            // -------------------------------------------------
            // Example 3: Remove a module by name (if it exists)
            // -------------------------------------------------
            string moduleToRemove = "OldModule";
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                if (vbaProject.Modules[i].Name.Equals(moduleToRemove, StringComparison.OrdinalIgnoreCase))
                {
                    vbaProject.Modules.RemoveAt(i);
                    break;
                }
            }

            // -------------------------------------------------
            // Save the modified workbook back to a macro‑enabled format
            // -------------------------------------------------
            string outputPath = "ModifiedWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with updated VBA modules to: {outputPath}");
        }
    }
}