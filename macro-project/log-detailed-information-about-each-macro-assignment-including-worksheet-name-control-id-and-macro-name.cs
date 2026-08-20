// Title: Log macro assignments (worksheet, control ID, macro name) with Aspose.Cells for .NET
// Description: C# sample that creates a macro‑enabled workbook, adds rectangle and oval shapes with assigned macros, then iterates every worksheet and shape to print the worksheet name, shape Name (used as control ID), and MacroName. The workbook is saved as an .xlsm file.
// Keywords: Aspose.Cells | C# macro logging | shape MacroName | worksheet name retrieval | control ID Excel | macro‑enabled workbook | iterate shapes Aspose.Cells | log Excel macros
// Common Searches: Aspose.Cells log shape macros | how to get macro name from shape in .NET | list all macro‑linked controls in Excel workbook | retrieve worksheet and shape IDs with Aspose.Cells | save macro‑enabled file after logging assignments
// Developer Intent: Extract and display each shape’s macro assignment together with its worksheet and control identifier.
// Use Cases: Generate a console audit of every macro‑linked control before publishing the workbook. | Validate that interactive shapes on all sheets have associated macros for quality checks. | Create a compliance log of worksheet, control ID, and macro name for downstream reporting.
// AI Prompts: Write C# code using Aspose.Cells to enumerate all worksheets and shapes, outputting worksheet name, shape Name, and MacroName. | Show how to filter out shapes without a MacroName and log only those with assigned macros. | Provide an example that saves the workbook as .xlsm after logging macro assignments, ensuring EnableMacros is true.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroLogging
{
    // C# sample that creates a macro‑enabled workbook, adds rectangle and oval shapes with assigned macros, then iterates every worksheet and shape to print the worksheet name, shape Name (used as control ID), and MacroName. The workbook is saved as an .xlsm file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable macros in the workbook settings (optional but good practice)
            workbook.Settings.EnableMacros = true;

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Add a rectangle shape and assign a macro to it
            Shape shape1 = sheet.Shapes.AddRectangle(1, 1, 120, 30, 0, 0);
            shape1.Name = "btnCalculate";               // Control ID
            shape1.MacroName = "CalculateValues()";     // Macro assignment

            // Add another shape with a macro
            Shape shape2 = sheet.Shapes.AddOval(5, 5, 100, 40, 0, 0);
            shape2.Name = "btnExport";
            shape2.MacroName = "ExportData()";

            // Log detailed information about each macro assignment
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Shape shp in ws.Shapes)
                {
                    if (!string.IsNullOrEmpty(shp.MacroName))
                    {
                        Console.WriteLine($"Worksheet: {ws.Name}, Control ID: {shp.Name}, Macro: {shp.MacroName}");
                    }
                }
            }

            // Save the workbook (macro-enabled format)
            workbook.Save("MacroAssignmentLogDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}
