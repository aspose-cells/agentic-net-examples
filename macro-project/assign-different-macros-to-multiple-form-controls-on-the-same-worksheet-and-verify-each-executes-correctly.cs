using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsMacroAssignment
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

            // -------------------------------------------------
            // Add first CommandButton ActiveX control and assign a macro
            // -------------------------------------------------
            Shape shape1 = sheet.Shapes.AddActiveXControl(
                ControlType.CommandButton, // control type
                2,   // upper left row
                2,   // upper left column
                2,   // top offset (pixels)
                2,   // left offset (pixels)
                120, // width (pixels)
                30   // height (pixels)
            );

            // Cast to CommandButtonActiveXControl to access specific properties if needed
            CommandButtonActiveXControl button1 = (CommandButtonActiveXControl)shape1.ActiveXControl;
            button1.Caption = "Run Macro One";
            // Assign macro name (the macro must exist in the VBA project; here we just set the name)
            shape1.MacroName = "MacroOne";

            // -------------------------------------------------
            // Add second CommandButton ActiveX control and assign a different macro
            // -------------------------------------------------
            Shape shape2 = sheet.Shapes.AddActiveXControl(
                ControlType.CommandButton,
                5,   // upper left row
                2,   // upper left column
                5,   // top offset
                2,   // left offset
                120,
                30
            );

            CommandButtonActiveXControl button2 = (CommandButtonActiveXControl)shape2.ActiveXControl;
            button2.Caption = "Run Macro Two";
            shape2.MacroName = "MacroTwo";

            // -------------------------------------------------
            // Verification: read back the assigned macro names and output to console
            // -------------------------------------------------
            Console.WriteLine("First button macro: " + shape1.MacroName);
            Console.WriteLine("Second button macro: " + shape2.MacroName);

            // Save the workbook as a macro‑enabled file
            workbook.Save("MultipleMacrosDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}