// Title: List ActiveX Form Controls in a Macro‑Enabled Excel Workbook with Aspose.Cells for .NET
// Description: Loads an .xlsm file using Aspose.Cells, walks through every worksheet and its shapes, detects ActiveX controls, and prints each control's worksheet, name and type to the console.
// Keywords: Aspose.Cells | C# | ActiveX controls | Excel form controls | macro‑enabled workbook | list controls | VBA shapes | enumerate ActiveX | xlsm processing | Excel automation
// Common Searches: how to list ActiveX controls in an .xlsm with Aspose.Cells | C# code to enumerate form controls that have macros assigned | retrieve control names and types from a macro‑enabled Excel file | Aspose.Cells example for reading VBA controls | list Excel ActiveX controls using .NET
// Developer Intent: The developer needs to load a macro‑enabled workbook and output all form controls that are linked to VBA macros.
// Use Cases: Create an inventory of VBA‑driven controls for documentation purposes. | Verify that required ActiveX controls exist before running further processing. | Generate an audit report of controls and their associated macros for compliance.
// AI Prompts: Generate C# code with Aspose.Cells that lists only ActiveX controls having a non‑empty OnAction macro reference. | Modify the sample to export the control list to a CSV file with columns: Worksheet, ControlName, ControlType, MacroName. | Add comprehensive error handling for missing VBA projects, corrupted shapes, and permission issues when opening .xlsm files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsVbaControlLister
{
    // Loads an .xlsm file using Aspose.Cells, walks through every worksheet and its shapes, detects ActiveX controls, and prints each control's worksheet, name and type to the console.
    class Program
    {
        static void Main()
        {
            const string inputFile = "input_with_macro.xlsm";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            try
            {
                // Load the workbook (no specific data filter needed for this example)
                Workbook workbook = new Workbook(inputFile);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape hosts an ActiveX control
                        if (shape.ActiveXControl != null)
                        {
                            Console.WriteLine($"Worksheet: {sheet.Name}");
                            Console.WriteLine($"  Control Name : {shape.Name}");
                            Console.WriteLine($"  Control Type : {shape.ActiveXControl.GetType().Name}");
                            Console.WriteLine();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
