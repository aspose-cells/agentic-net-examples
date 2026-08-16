// Title: Assign Different Macros to Multiple ActiveX Command Buttons with Aspose.Cells for .NET
// Description: Demonstrates how to create a macro‑enabled workbook, add two ActiveX CommandButton controls, set individual captions, assign distinct MacroName values (Macro1, Macro2), verify the assignments via console output, and save the file as an .xlsm workbook using Aspose.Cells.
// Keywords: Aspose.Cells assign macro | ActiveX CommandButton macro name | macro‑enabled workbook .NET | set MacroName property | multiple form controls Aspose.Cells | C# Excel automation
// Common Searches: how to link different macros to ActiveX buttons in Aspose.Cells | verify MacroName of ActiveX controls C# | save .xlsm with multiple command buttons using Aspose.Cells | assign macro to ActiveX CommandButton programmatically
// Developer Intent: Programmatically attach unique macro identifiers to several ActiveX command buttons on the same worksheet and confirm each assignment.
// Use Cases: Generate a template workbook where each button launches a specific VBA routine. | Automate validation of macro links before distributing macro‑enabled Excel files. | Create interactive dashboards with pre‑wired controls for end‑user actions.
// AI Prompts: Write C# code with Aspose.Cells to add three ActiveX checkboxes, assign separate macro names, and print each name for verification. | Explain how to read, modify, and persist the MacroName of an existing ActiveX control in a saved .xlsm file using Aspose.Cells. | Provide a step‑by‑step guide to create a macro‑enabled workbook, link multiple form controls to different macros, and implement unit‑test style checks for the assignments.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Demonstrates how to create a macro‑enabled workbook, add two ActiveX CommandButton controls, set individual captions, assign distinct MacroName values (Macro1, Macro2), verify the assignments via console output, and save the file as an .xlsm workbook using Aspose.Cells.
class AssignMacrosDemo
{
    static void Main()
    {
        // Create a new workbook and enable macros
        Workbook workbook = new Workbook();
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Add first CommandButton ActiveX control
        // -------------------------------------------------
        Shape shape1 = sheet.Shapes.AddActiveXControl(
            ControlType.CommandButton, // control type
            1, 0,                     // upper left row, column
            1, 0,                     // lower right row, column
            100, 30);                 // width, height in pixels

        // Cast to specific control to set its properties
        CommandButtonActiveXControl button1 = (CommandButtonActiveXControl)shape1.ActiveXControl;
        button1.Caption = "Run Macro1";
        // Assign macro name to the shape
        shape1.MacroName = "Macro1";

        // -------------------------------------------------
        // Add second CommandButton ActiveX control
        // -------------------------------------------------
        Shape shape2 = sheet.Shapes.AddActiveXControl(
            ControlType.CommandButton,
            3, 0,
            1, 0,
            100, 30);

        CommandButtonActiveXControl button2 = (CommandButtonActiveXControl)shape2.ActiveXControl;
        button2.Caption = "Run Macro2";
        shape2.MacroName = "Macro2";

        // -------------------------------------------------
        // Verify that each shape has the correct macro name
        // -------------------------------------------------
        Console.WriteLine("Button 1 MacroName: " + shape1.MacroName);
        Console.WriteLine("Button 2 MacroName: " + shape2.MacroName);

        // Save the workbook as a macro‑enabled file
        workbook.Save("MultipleMacrosDemo.xlsm");
    }
}
