using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroLinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSM will be used when saving)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Add VBA macro to the workbook
            // -------------------------------------------------
            // Ensure the workbook has a macro project by saving as XLSM and reloading if necessary
            if (workbook.VbaProject == null)
            {
                // Save as a temporary macro-enabled file to initialise the VBA project
                string tempPath = "temp.xlsm";
                workbook.Save(tempPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempPath);
                System.IO.File.Delete(tempPath);
            }

            // Add a procedural module and insert macro code
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MacroModule");
            VbaModule macroModule = workbook.VbaProject.Modules[moduleIndex];
            // Macro name must match the name we will assign to the button later
            string macroName = "MyMacro";
            macroModule.Codes = $"Sub {macroName}()\n    MsgBox \"Button clicked!\"\nEnd Sub";

            // -------------------------------------------------
            // 2. Add a Forms button control to the first worksheet
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button shape (Forms control). Parameters: upper left row, upper left column,
            // top offset (pixels), left offset (pixels), width (pixels), height (pixels)
            Shape buttonShape = sheet.Shapes.AddButton(2, 2, 0, 0, 120, 30);
            // Set the visible caption of the button
            buttonShape.Text = "Click Me";

            // Link the button to the VBA macro using the MacroName property
            buttonShape.MacroName = $"{macroName}()";

            // -------------------------------------------------
            // 3. Save the workbook as a macro‑enabled file
            // -------------------------------------------------
            string outputPath = "LinkedMacroButton.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to '{outputPath}'. The button is linked to macro '{macroName}'.");
        }
    }
}