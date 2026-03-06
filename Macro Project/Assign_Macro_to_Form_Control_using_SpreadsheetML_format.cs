using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable macros for the workbook
            workbook.Settings.EnableMacros = true;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button control to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Shape button = sheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

            // Assign a macro name to the button (the macro must exist in the VBA project)
            button.MacroName = "MyMacro()";

            // Save the workbook in SpreadsheetML (Excel 2003 XML) format
            workbook.Save("MacroButton.xml", SaveFormat.SpreadsheetML);
        }
    }
}