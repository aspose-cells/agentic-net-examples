using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsButtonVbaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button to the worksheet (row 1, column 1, size 100x30 pixels)
            Button button = sheet.Shapes.AddButton(0, 0, 0, 0, 100, 30);
            button.Text = "Highlight Row";
            button.Name = "HighlightButton";

            // Assign the macro name that will be executed when the button is clicked
            button.MacroName = "HighlightActiveRow";

            // Ensure the workbook has a VBA project (required for macro-enabled files)
            // Adding a module for the worksheet and inserting VBA code
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "StandardModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];

            // VBA macro: highlights the entire active row with yellow background
            string vbaCode = @"
Sub HighlightActiveRow()
    Dim r As Long
    r = ActiveCell.Row
    Rows(r).Interior.Color = RGB(255, 255, 0)
End Sub
";
            vbaModule.Codes = vbaCode.Trim();

            // Save the workbook as a macro-enabled file
            workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}