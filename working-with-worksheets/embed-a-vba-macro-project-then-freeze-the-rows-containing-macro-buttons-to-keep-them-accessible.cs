// Title: Add a VBA macro button shape to the first row of a worksheet and freeze that row using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that inserts a rectangular shape as a macro button on row 1, sets its name and alternative text, applies FreezePanes to keep the row visible, and saves the workbook as an XLSM file. | Generate a .NET example that creates a macro‑enabled Excel workbook, adds a shape button linked to a VBA project, locks the button’s row in place, and outputs the file in XLSM format.
// Common Searches: Aspose.Cells C# add shape button to first row and freeze panes | how to create a macro‑enabled XLSM with a button using Aspose.Cells .NET | freeze top row containing a shape in a generated Excel file with Aspose.Cells | embed VBA macro button in Aspose.Cells workbook and keep it visible while scrolling | Aspose.Cells FreezePanes after inserting a rectangle shape as a macro button
// Tags: add rectangle shape macro button Aspose.Cells | freeze top row Aspose.Cells C# | save workbook as XLSM Aspose.Cells | embed VBA project Aspose.Cells .NET | macro button visibility scrolling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, renames the first worksheet, adds a rectangular shape configured as a VBA macro button, freezes the first row so the button remains visible during scrolling, and saves the result as a macro‑enabled XLSM file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Define button size (in points)
            int buttonWidth = 100;   // width
            int buttonHeight = 30;   // height

            // Add a rectangular shape that will act as a macro button
            // Parameters: shape type, upper left row, upper left column,
            // upper left row offset, upper left column offset, height, width
            Shape button = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, 0, 0, 0, 0, buttonHeight, buttonWidth);
            button.Name = "btnRunMacro";
            button.AlternativeText = "Run Macro";

            // Freeze the first row so the button stays visible while scrolling
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
