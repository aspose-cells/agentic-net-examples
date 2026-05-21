using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsMacroAssignmentDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook();

            // Enable macros for the workbook
            workbook.Settings.EnableMacros = true;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------- Add first ActiveX CommandButton --------------------
            // Parameters: control type, upper left row, upper left column, top offset, left offset, width, height
            Shape cmdButtonShape1 = sheet.Shapes.AddActiveXControl(
                ControlType.CommandButton, 2, 2, 0, 0, 120, 30);
            CommandButtonActiveXControl cmdButton1 = (CommandButtonActiveXControl)cmdButtonShape1.ActiveXControl;
            cmdButton1.Caption = "Run Macro A";
            // Assign macro name to the shape (not the control)
            cmdButtonShape1.MacroName = "MacroA";

            // -------------------- Add second ActiveX CommandButton --------------------
            Shape cmdButtonShape2 = sheet.Shapes.AddActiveXControl(
                ControlType.CommandButton, 5, 2, 0, 0, 120, 30);
            CommandButtonActiveXControl cmdButton2 = (CommandButtonActiveXControl)cmdButtonShape2.ActiveXControl;
            cmdButton2.Caption = "Run Macro B";
            cmdButtonShape2.MacroName = "MacroB";

            // -------------------- Add a Forms Button (inherits Shape) --------------------
            Shape formsButton = sheet.Shapes.AddButton(8, 2, 0, 0, 120, 30);
            formsButton.Text = "Run Macro C";
            formsButton.MacroName = "MacroC";

            // -------------------- Verify macro assignments (in‑memory) --------------------
            Console.WriteLine("Macro assigned to first CommandButton: " + cmdButtonShape1.MacroName);
            Console.WriteLine("Macro assigned to second CommandButton: " + cmdButtonShape2.MacroName);
            Console.WriteLine("Macro assigned to Forms Button: " + formsButton.MacroName);

            // -------------------- Save the workbook --------------------
            string filePath = "MultipleMacrosDemo.xlsm";
            workbook.Save(filePath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to {filePath}");

            // -------------------- Load the workbook back and re‑verify --------------------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedSheet = loadedWb.Worksheets[0];

            // Retrieve shapes by index (order of addition)
            Shape loadedCmdButtonShape1 = loadedSheet.Shapes[0];
            Shape loadedCmdButtonShape2 = loadedSheet.Shapes[1];
            Shape loadedFormsButton = loadedSheet.Shapes[2];

            Console.WriteLine("After reload - Macro of first CommandButton: " + loadedCmdButtonShape1.MacroName);
            Console.WriteLine("After reload - Macro of second CommandButton: " + loadedCmdButtonShape2.MacroName);
            Console.WriteLine("After reload - Macro of Forms Button: " + loadedFormsButton.MacroName);
        }
    }
}