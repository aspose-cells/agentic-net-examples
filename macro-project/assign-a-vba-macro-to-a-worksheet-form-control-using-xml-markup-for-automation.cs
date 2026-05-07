using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroAssignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add a Forms button (Button shape) to the sheet
            // -------------------------------------------------
            // Parameters: upper left row, upper left column, upper left row offset,
            // upper left column offset, lower right row, lower right column
            Button btn = (Button)sheet.Shapes.AddButton(2, 2, 0, 0, 4, 4);
            btn.Text = "Run Macro";

            // -------------------------------------------------
            // 2. Create a VBA module and add macro code
            // -------------------------------------------------
            // Ensure the workbook is macro‑enabled by saving as .xlsm later
            VbaProject vbaProject = workbook.VbaProject;
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacros");
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];
            vbaModule.Codes = @"
Sub MyMacro()
    MsgBox ""Hello from the assigned macro!""
End Sub
";

            // -------------------------------------------------
            // 3. Assign the macro to the button using Shape.MacroName
            // -------------------------------------------------
            // The macro name must include the module name (optional) and the sub name
            btn.MacroName = "MyMacro";

            // -------------------------------------------------
            // 4. Save the workbook as a macro‑enabled file
            // -------------------------------------------------
            workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}