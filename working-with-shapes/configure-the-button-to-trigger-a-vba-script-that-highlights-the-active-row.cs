// Title: Create a Forms button in an Excel .xlsm file and bind it to a VBA macro that highlights the active row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a Forms button that runs a VBA macro to highlight the selected row | How to assign a VBA macro to a button in a macro‑enabled workbook with Aspose.Cells | C# generate a macro‑enabled .xlsm, insert a button, and link it to HighlightActiveRow macro | Using Aspose.Cells for .NET to create an Excel button that colors the active row via VBA | Programmatically bind a VBA sub to an Excel button to shade the active row with Aspose.Cells
// Tags: forms button insertion Aspose.Cells | vba macro assignment to Excel button C# | active row highlighting VBA Aspose.Cells | macro enabled xlsm generation Aspose.Cells | button click macro binding .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Drawing;

namespace AsposeCellsButtonVbaDemo
{
    // The C# example creates a macro‑enabled .xlsm workbook, adds a Forms button at B2, defines a VBA macro that fills the active row with yellow, links the button to this macro, and saves the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path where the workbook will be saved
                const string outputPath = "ButtonHighlightActiveRow.xlsm";

                // Create a new workbook (empty)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // ------------------------------------------------------------
                // 1. Add a Forms button to the worksheet
                // ------------------------------------------------------------
                // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (px), width (px)
                // Place the button at cell B2 (row index 1, column index 1)
                Button button = sheet.Shapes.AddButton(
                    topRow: 1,          // Upper‑left row index (zero‑based)
                    top: 0,             // Pixel offset from the top of the row
                    leftColumn: 1,      // Upper‑left column index (zero‑based)
                    left: 0,            // Pixel offset from the left of the column
                    height: 30,         // Height in pixels
                    width: 120);        // Width in pixels

                // Configure button appearance
                button.Name = "HighlightRowButton";
                button.Text = "Highlight Row";
                button.Font.Size = 10;
                button.Font.IsBold = true;                     // Use IsBold instead of Bold
                button.Placement = PlacementType.FreeFloating; // Do not move/size with cells

                // ------------------------------------------------------------
                // 2. Create a VBA macro that highlights the active row
                // ------------------------------------------------------------
                // Ensure the workbook has a VBA project (it is created automatically)
                VbaProject vbaProject = workbook.VbaProject;
                vbaProject.Name = "ButtonVbaProject";

                // Add a standard module to hold the macro
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "HighlightModule"); // Use Class type
                VbaModule vbaModule = vbaProject.Modules[moduleIndex];

                // VBA code: highlight the entire row of the active cell with yellow background
                string macroCode = @"
Sub HighlightActiveRow()
    Dim rng As Range
    Set rng = ActiveCell.EntireRow
    rng.Interior.ColorIndex = 6   ' Yellow (standard palette index)
End Sub
";

                vbaModule.Codes = macroCode.Trim();

                // ------------------------------------------------------------
                // 3. Link the button to the macro
                // ------------------------------------------------------------
                // The MacroName property expects the macro name (without module prefix)
                button.MacroName = "HighlightActiveRow";

                // ------------------------------------------------------------
                // 4. Save the workbook as a macro‑enabled file
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Open the file in Excel, click the button, and the active row will be highlighted.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
