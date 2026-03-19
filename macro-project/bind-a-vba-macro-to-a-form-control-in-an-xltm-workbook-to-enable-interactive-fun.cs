using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Drawing;

class BindMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook (will be saved as a macro‑enabled template later)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural VBA module to hold the macro code
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MacroModule");
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Define a simple macro that shows a message box
        string macroCode =
            "Sub MyButtonMacro()\r\n" +
            "    MsgBox \"Button clicked!\"\r\n" +
            "End Sub";

        // Assign the code to the newly created module
        module.Codes = macroCode;

        // Get the first worksheet where the form control will be placed
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will serve as a button (Form control)
        // Parameters: upper left row, upper left column, width, height, upper left offset X, offset Y
        Shape button = sheet.Shapes.AddRectangle(2, 2, 100, 30, 0, 0);

        // Set the display text of the button
        button.Text = "Click Me";

        // Bind the VBA macro to the shape using the MacroName property
        // The macro name must match the Sub name defined above
        button.MacroName = "MyButtonMacro";

        // Save the workbook as a macro‑enabled template (.xltm)
        workbook.Save("MacroBoundTemplate.xltm", SaveFormat.Xltm);
    }
}