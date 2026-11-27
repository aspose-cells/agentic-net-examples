// Title: Embedding a VBA macro project and freezing rows with macro buttons using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed VBA project and add a macro button | freeze panes on rows that contain VBA button controls with Aspose.Cells | create macro‑enabled XLSM workbook and keep button rows visible in .NET | determine button rows and apply FreezePanes in Aspose.Cells | assign macro to Forms button and freeze top rows in C#
// Tags: VBA project embedding with Aspose.Cells | inserting Forms button linked to macro Aspose.Cells | freezing top rows based on button location Aspose.Cells | saving as macro‑enabled XLSM using Aspose.Cells | detecting macro button rows C# Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Drawing;

namespace AsposeCellsVbaFreezeRows
{
    // The example creates a new workbook, fills it with sample data, embeds a VBA module containing a simple macro, adds a Forms button that triggers the macro, identifies the rows where macro buttons reside, freezes panes up to the lowest button row to keep the buttons visible, and saves the result as a macro‑enabled XLSM file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // 2. Populate sample data (rows 1‑5)
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(2.5);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(1.8);
                cells["A4"].PutValue("Banana");
                cells["B4"].PutValue(1.2);
                cells["A5"].PutValue("Grapes");
                cells["B5"].PutValue(3.0);

                // 3. Add a VBA module with a simple macro
                VbaProject vbaProject = workbook.VbaProject;
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "MacroModule");
                VbaModule vbaModule = vbaProject.Modules[moduleIndex];
                vbaModule.Codes =
@"Sub ShowMessage()
    MsgBox ""Hello from VBA macro!""
End Sub";

                // 4. Insert a button shape that will trigger the macro
                // Parameters: upper left row, column, upper left pixel row, column, height (points), width (points)
                int buttonRow = 2;          // zero‑based row (Excel row 3)
                int buttonColumn = 2;       // column C
                int buttonHeight = 30;      // points
                int buttonWidth = 100;      // points

                // AddButton returns a Button object (Forms control)
                Button button = sheet.Shapes.AddButton(buttonRow, buttonColumn, 0, 0, buttonHeight, buttonWidth);
                button.Name = "ShowMsgButton";
                button.AlternativeText = "Click to show message";
                // Assign the macro to the button
                button.MacroName = "MacroModule.ShowMessage";

                // 5. Determine rows that contain macro buttons
                List<int> buttonRows = new List<int>();
                foreach (Shape s in sheet.Shapes)
                {
                    if (s is Button btn && !string.IsNullOrEmpty(btn.MacroName))
                    {
                        buttonRows.Add(btn.UpperLeftRow);
                    }
                }

                // 6. Freeze rows up to the lowest button row (inclusive)
                if (buttonRows.Count > 0)
                {
                    int maxButtonRow = buttonRows.Max();
                    // Freeze panes just below the max button row
                    sheet.FreezePanes(maxButtonRow + 1, 0, maxButtonRow + 1, 0);
                    Console.WriteLine($"Frozen rows 1 through {maxButtonRow + 1} to keep macro buttons visible.");
                }
                else
                {
                    Console.WriteLine("No macro buttons found; no rows were frozen.");
                }

                // 7. Save the workbook as a macro‑enabled file
                string outputPath = "VbaWithFrozenButtonRows.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
