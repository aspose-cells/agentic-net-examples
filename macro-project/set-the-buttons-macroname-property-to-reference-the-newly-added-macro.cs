using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroButtonDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable macros in the workbook
            workbook.Settings.EnableMacros = true;

            // Add a VBA module to the workbook and insert a simple macro
            int moduleIndex = workbook.VbaProject.Modules.Add(sheet);
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Name = "Module1";
            module.Codes = @"
Sub MyMacro()
    MsgBox ""Hello from MyMacro!""
End Sub";

            // Add a Forms button to the worksheet
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
            Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
            button.Text = "Run Macro";

            // Associate the button with the newly added macro
            button.MacroName = "MyMacro";

            // Save the workbook as a macro‑enabled file
            workbook.Save("ButtonWithMacro.xlsm");
        }
    }
}