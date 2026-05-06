using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroFormControlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (initially without a VBA project)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project by saving as a macro‑enabled file and reloading it
            string tempMacroFile = "temp.xlsm";
            workbook.Save(tempMacroFile, SaveFormat.Xlsm);
            workbook = new Workbook(tempMacroFile);
            File.Delete(tempMacroFile); // clean up the temporary file

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as a form control (e.g., a button)
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, lower right row, lower right column
            Shape buttonShape = sheet.Shapes.AddRectangle(2, 2, 30, 100, 2, 2);

            // Set the macro name that will be executed when the shape is clicked
            buttonShape.MacroName = "ButtonClick";

            // Add a procedural VBA module to hold the macro code
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MacroModule");
            VbaModule macroModule = workbook.VbaProject.Modules[moduleIndex];

            // Define the macro code (must match the MacroName set on the shape)
            macroModule.Codes =
                "Sub ButtonClick()\n" +
                "    MsgBox \"Button clicked!\"\n" +
                "End Sub";

            // Save the workbook as a macro‑enabled file (XLSM)
            string outputPath = "WorkbookWithMacroAndFormControl.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Macro is linked to the shape.");
        }
    }
}