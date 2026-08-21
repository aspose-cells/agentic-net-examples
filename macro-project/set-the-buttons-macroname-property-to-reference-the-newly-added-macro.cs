// Title: Assign a VBA macro to a Forms button in an Excel workbook using Aspose.Cells C#
// Description: Demonstrates how to create a macro‑enabled .xlsm workbook with Aspose.Cells, add a VBA module containing a Sub, insert a Forms button, set the button's MacroName property to the new macro, and save the file.
// Keywords: Aspose.Cells C# macro button | set button MacroName Aspose.Cells | add VBA module Aspose.Cells | macro‑enabled workbook .xlsm | Forms button VBA Aspose.Cells | C# Excel automation Aspose | link VBA macro to button
// Common Searches: How to link a VBA macro to a Forms button with Aspose.Cells .NET | Set MacroName property for a button in an .xlsm file using C# | Create macro‑enabled workbook and assign macro to button Aspose.Cells | Aspose.Cells example for adding VBA module and button
// Developer Intent: Connect a newly created VBA macro to a Forms button so that clicking the button runs the macro in the saved workbook.
// Use Cases: Generate an Excel template that includes a button to display a message box via VBA. | Automate the creation of multiple buttons, each bound to a distinct macro, for complex worksheets. | Provide end‑users with a clickable control that triggers custom VBA logic without manual macro insertion.
// AI Prompts: Write C# code with Aspose.Cells that adds a VBA module, defines a Sub, and assigns it to a Forms button's MacroName property. | Explain the steps to enable macros, add a button, link it to a macro, and save the workbook as .xlsm using Aspose.Cells. | Troubleshoot why a button's MacroName does not execute the macro after the workbook is opened.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a macro‑enabled .xlsm workbook with Aspose.Cells, add a VBA module containing a Sub, insert a Forms button, set the button's MacroName property to the new macro, and save the file.
    public class SetButtonMacroDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable macros in the workbook
                workbook.Settings.EnableMacros = true;

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a VBA module to the workbook and insert a macro
                int moduleIndex = workbook.VbaProject.Modules.Add(sheet);
                VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
                vbaModule.Name = "Module1";
                vbaModule.Codes =
                    "Sub MyMacro()\r\n" +
                    "    MsgBox \"Hello from macro\"\r\n" +
                    "End Sub";

                // Add a Forms button to the worksheet
                // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixel), width (pixel)
                Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
                button.Text = "Run Macro";

                // Associate the button with the newly added macro
                button.MacroName = "MyMacro";

                // Save the workbook as a macro‑enabled file
                string outputPath = "ButtonWithMacro.xlsm";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetButtonMacroDemo.Run();
        }
    }
}
