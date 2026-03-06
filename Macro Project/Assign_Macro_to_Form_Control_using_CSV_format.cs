using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroFromCsv
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button control (row, column, top, left, height, width)
        Button button = worksheet.Shapes.AddButton(1, 1, 0, 0, 30, 100);
        button.Name = "Button1";

        // CSV file format: ShapeName,MacroName
        // Example line: Button1,MyMacro()
        string csvFilePath = "macros.csv";

        if (File.Exists(csvFilePath))
        {
            foreach (string line in File.ReadAllLines(csvFilePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue; // Skip empty lines

                // Split by comma
                string[] parts = line.Split(',');
                if (parts.Length < 2)
                    continue; // Invalid line

                string shapeName = parts[0].Trim();
                string macroName = parts[1].Trim();

                // Locate the shape by name
                Shape shape = worksheet.Shapes[shapeName];
                if (shape != null)
                {
                    // Assign the macro name to the shape
                    shape.MacroName = macroName;
                }
            }
        }

        // Save the workbook as a macro‑enabled file
        workbook.Save("Result.xlsm", SaveFormat.Xlsm);
    }
}