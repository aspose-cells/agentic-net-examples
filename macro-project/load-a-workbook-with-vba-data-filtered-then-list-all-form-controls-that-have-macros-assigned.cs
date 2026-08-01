// Title: C# – List ActiveX Form Controls with Assigned Macros in a Macro‑Enabled Workbook Using Aspose.Cells
// Description: Loads an XLSM file with Aspose.Cells, confirms the presence of VBA macros, then walks through each worksheet and its shapes to identify ActiveX controls. For every control it prints the worksheet name, shape name, control type and, when available, the linked cell reference.
// Keywords: Aspose.Cells | C# | .NET | XLSM | macro enabled workbook | VBA macros | ActiveX control | form control listing | enumerate shapes | shape.ActiveXControl | workbook.HasMacro | linked cell | code example
// Common Searches: list ActiveX controls in an .xlsm with Aspose.Cells | how to enumerate VBA form controls using C# | retrieve linked cell of ActiveX control Aspose.Cells | C# code to read macro enabled workbook shapes | Aspose.Cells get form control properties
// Developer Intent: Extract all ActiveX form controls from a macro‑enabled workbook and display their key attributes.
// Use Cases: Create an inventory of form controls before refactoring or removing macros. | Generate documentation of control types and linked cells for compliance audits. | Automated testing to verify required controls exist in a workbook.
// AI Prompts: Show C# code that uses Aspose.Cells to enumerate every ActiveX control in an XLSM file and output its type and linked cell. | Explain how to filter shapes to only those that have VBA macros attached when processing a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsVbaFormControlListing
{
    // Loads an XLSM file with Aspose.Cells, confirms the presence of VBA macros, then walks through each worksheet and its shapes to identify ActiveX controls. For every control it prints the worksheet name, shape name, control type and, when available, the linked cell reference.
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook
            string inputPath = "SampleWithMacro.xlsm";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (full load; VBA data will be available if present)
                Workbook workbook = new Workbook(inputPath);

                // Verify that the workbook actually contains VBA macros
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any VBA macros.");
                    return;
                }

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape hosts an ActiveX control
                        if (shape.ActiveXControl != null)
                        {
                            Console.WriteLine($"Worksheet: {sheet.Name}");
                            Console.WriteLine($"  Shape Name   : {shape.Name}");
                            Console.WriteLine($"  Control Type : {shape.ActiveXControl.GetType().Name}");
                            // Additional useful information (e.g., linked cell) can be displayed if needed
                            if (!string.IsNullOrEmpty(shape.ActiveXControl.LinkedCell))
                            {
                                Console.WriteLine($"  Linked Cell  : {shape.ActiveXControl.LinkedCell}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
