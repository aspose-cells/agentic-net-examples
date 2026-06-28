using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class MacroLogger
{
    static void Main()
    {
        // Load a macro‑enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each shape on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // If the shape has a macro assigned, log the details
                if (!string.IsNullOrEmpty(shape.MacroName))
                {
                    // Worksheet name, shape (control) ID, and macro name
                    Console.WriteLine($"Worksheet: {sheet.Name}, Control ID: {shape.Id}, Macro Name: {shape.MacroName}");
                }
            }
        }

        // Save the workbook (unchanged) – optional
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}