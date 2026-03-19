using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the XLTX template (must exist)
            string templatePath = "Template.xltx";

            // Load the XLTX workbook
            Workbook workbook = new Workbook(templatePath);

            // Ensure the workbook has a VBA project.
            // If not, save as a macro‑enabled file and reload.
            if (workbook.VbaProject == null)
            {
                string tempMacroPath = "temp.xlsm";
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);
                workbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);
            }

            // Add a procedural VBA module and insert macro code.
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MacroModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Codes =
                "Sub MyButtonClick()\n" +
                "    MsgBox \"Button clicked!\"\n" +
                "End Sub";

            // Add a shape (rectangle) that will act as the form control.
            Worksheet sheet = workbook.Worksheets[0];
            // Parameters: upper left row, upper left column, height (pixels), width (pixels), top offset, left offset
            Shape buttonShape = sheet.Shapes.AddRectangle(1, 1, 30, 100, 0, 0);
            buttonShape.Name = "MyButton";

            // Associate the VBA macro with the shape.
            buttonShape.MacroName = "MyButtonClick";

            // Save the workbook as a macro‑enabled file.
            string outputPath = "WorkbookWithMacro.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to '{outputPath}' with macro linked to the shape.");
        }
    }
}