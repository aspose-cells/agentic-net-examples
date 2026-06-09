using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace MacroCopyExample
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the macro "InitializeReport"
            Workbook templateWorkbook = new Workbook("TemplateWithMacro.xlsm");

            // Load (or create) the target workbook where the macro will be copied to
            Workbook targetWorkbook = new Workbook();

            // Ensure the target workbook can hold macros
            targetWorkbook.Settings.EnableMacros = true;

            // Retrieve the designer storage (binary data) of the macro named "InitializeReport"
            // This works for macros stored as Designer modules (e.g., userforms or designer code)
            byte[] macroStorage = templateWorkbook.VbaProject.Modules.GetDesignerStorage("InitializeReport");

            if (macroStorage != null && macroStorage.Length > 0)
            {
                // Add the retrieved designer storage to the target workbook's VBA project
                // The AddDesignerStorage method creates a Designer module with the same name
                targetWorkbook.VbaProject.Modules.AddDesignerStorage("InitializeReport", macroStorage);
                Console.WriteLine("Macro 'InitializeReport' copied successfully via DesignerStorage.");
            }
            else
            {
                Console.WriteLine("Designer storage for macro 'InitializeReport' not found in the template workbook.");
            }

            // Save the target workbook as a macro‑enabled file
            targetWorkbook.Save("TargetWithCopiedMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}