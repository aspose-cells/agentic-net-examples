using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled will be saved later)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project.
            // If the workbook does not contain a macro, save as .xlsm and reload to create the project.
            if (workbook.VbaProject == null)
            {
                workbook.Save("temp.xlsm", SaveFormat.Xlsm);
                workbook = new Workbook("temp.xlsm");
            }

            // Add a procedural VBA module and insert a simple macro.
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
            vbaModule.Codes =
                "Sub MyMacro()\r\n" +
                "    MsgBox \"Macro triggered from shape!\"\r\n" +
                "End Sub";

            // Add a rectangle shape that will act as a button.
            Worksheet sheet = workbook.Worksheets[0];
            Shape buttonShape = sheet.Shapes.AddRectangle(2, 2, 100, 30, 0, 0);

            // Assign the macro to the shape.
            buttonShape.MacroName = "MyMacro";

            // Optionally set a caption for visual clarity.
            buttonShape.Text = "Run Macro";

            // Save the workbook as a macro‑enabled file.
            workbook.Save("WorkbookWithMacroAndShape.xlsm", SaveFormat.Xlsm);
        }
    }
}