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
            // Create a new workbook (initially without a VBA project)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project.
            // If VbaProject is null, save as a macro‑enabled file and reload.
            if (workbook.VbaProject == null)
            {
                string tempXlsm = "temp.xlsm";
                workbook.Save(tempXlsm, SaveFormat.Xlsm);
                workbook = new Workbook(tempXlsm);
                File.Delete(tempXlsm);
            }

            // Add a procedural VBA module and insert a simple macro.
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
            vbaModule.Codes =
                "Sub MyMacro()\n" +
                "    MsgBox \"Hello from VBA macro!\"\n" +
                "End Sub";

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a shape (button) to the worksheet.
            // Parameters: upper left row, upper left column, height, width, offsetX, offsetY
            Shape button = sheet.Shapes.AddRectangle(2, 2, 30, 100, 0, 0);
            button.Name = "MyButton";

            // Associate the VBA macro with the shape.
            button.MacroName = "MyMacro";

            // Save the workbook as ODS (OpenDocument Spreadsheet).
            workbook.Save("WorkbookWithMacro.ods", SaveFormat.Ods);
        }
    }
}