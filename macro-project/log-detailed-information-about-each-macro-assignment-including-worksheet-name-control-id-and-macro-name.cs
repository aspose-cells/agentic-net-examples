using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class MacroLogger
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable macros for the workbook (optional but recommended)
        workbook.Settings.EnableMacros = true;

        // ---------- Worksheet 1 ----------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "SheetOne";

        // Add a rectangle shape and assign a macro
        Shape rectShape = sheet1.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
        rectShape.Name = "Button1";          // Control ID
        rectShape.MacroName = "MacroOne";    // Macro name

        // ---------- Worksheet 2 ----------
        int newSheetIdx = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[newSheetIdx];
        sheet2.Name = "SheetTwo";

        // Add an oval shape and assign a macro
        Shape ovalShape = sheet2.Shapes.AddOval(2, 2, 80, 80, 0, 0);
        ovalShape.Name = "Button2";          // Control ID
        ovalShape.MacroName = "MacroTwo";    // Macro name

        // Log detailed information about each macro assignment
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Shape shp in ws.Shapes)
            {
                // Only log shapes that have a macro assigned
                if (!string.IsNullOrEmpty(shp.MacroName))
                {
                    Console.WriteLine($"Worksheet: {ws.Name}, Control ID: {shp.Name}, Macro Name: {shp.MacroName}");
                }
            }
        }

        // Save the workbook as a macro‑enabled file
        workbook.Save("MacroLogDemo.xlsm", SaveFormat.Xlsm);
    }
}