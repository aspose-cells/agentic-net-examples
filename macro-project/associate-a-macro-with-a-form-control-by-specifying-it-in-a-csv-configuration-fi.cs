using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsMacroAssociation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable macros in the workbook
            workbook.Settings.EnableMacros = true;

            // Add a CommandButton ActiveX control as an example
            Shape buttonShape = sheet.Shapes.AddActiveXControl(
                ControlType.CommandButton, // control type
                2, 2,                     // upper‑left row, column
                2, 2,                     // upper‑left row offset, column offset
                120, 30);                 // width, height

            // Assign a name to the shape so it can be referenced from CSV
            buttonShape.Name = "MyButton";

            // Path to the CSV configuration file
            // Expected format per line: ShapeName,MacroName
            string csvPath = "MacroConfig.csv";

            if (File.Exists(csvPath))
            {
                // Read all lines from the CSV
                string[] lines = File.ReadAllLines(csvPath);
                foreach (string line in lines)
                {
                    // Skip empty lines
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Split by comma
                    string[] parts = line.Split(',');
                    if (parts.Length != 2)
                        continue; // Invalid line format

                    string shapeName = parts[0].Trim();
                    string macroName = parts[1].Trim();

                    // Find the shape by name
                    Shape shape = null;
                    foreach (Shape s in sheet.Shapes)
                    {
                        if (s.Name == shapeName)
                        {
                            shape = s;
                            break;
                        }
                    }

                    if (shape != null)
                    {
                        // Associate the macro with the shape
                        shape.MacroName = macroName;
                    }
                }
            }
            else
            {
                Console.WriteLine($"CSV configuration file not found: {csvPath}");
            }

            // Save the workbook (macro‑enabled format)
            workbook.Save("WorkbookWithMacros.xlsm");
        }
    }
}