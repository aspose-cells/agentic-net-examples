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
            Workbook targetWorkbook = new Workbook(); // creates a new empty workbook
            // Ensure the target workbook is macro‑enabled
            targetWorkbook.Settings.EnableMacros = true;

            // Retrieve the designer storage (binary data) of the macro named "InitializeReport"
            // from the template workbook's VBA project.
            byte[] macroStorage = templateWorkbook.VbaProject.Modules.GetDesignerStorage("InitializeReport");

            if (macroStorage == null)
            {
                Console.WriteLine("Macro 'InitializeReport' not found in the template workbook.");
                return;
            }

            // Add the retrieved designer storage to the target workbook's VBA project.
            // This creates a new designer module with the same name and code.
            targetWorkbook.VbaProject.Modules.AddDesignerStorage("InitializeReport", macroStorage);

            // Save the target workbook as a macro‑enabled file.
            targetWorkbook.Save("TargetWithCopiedMacro.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Macro 'InitializeReport' successfully copied to the target workbook.");
        }
    }
}