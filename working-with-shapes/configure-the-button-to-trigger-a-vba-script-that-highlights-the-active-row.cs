// Title: C# – Add a Form Button that Runs a VBA Macro to Highlight the Active Row (Aspose.Cells)
// Description: Creates a new workbook, inserts a Form button on the first worksheet, assigns the macro name HighlightActiveRow, embeds a VBA module that colors the active cell's entire row yellow, and saves the file as a macro‑enabled .xlsm using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# button macro | add form button Excel .NET | VBA macro HighlightActiveRow | save workbook as xlsm | assign macro to shape Aspose
// Common Searches: how to add a button that runs a VBA macro with Aspose.Cells | C# code to create a macro‑enabled Excel file and attach a button | Aspose.Cells example: highlight active row via button | programmatically add Form control and VBA in .NET workbook
// Developer Intent: Generate a macro‑enabled Excel workbook, place a Form button, and link it to a VBA routine that highlights the row of the active cell.
// Use Cases: Provide end‑users a one‑click way to emphasize the row they are editing in a data‑entry template. | Automate report sheets where a button instantly marks the current record row for review. | Include a reusable highlight‑row button in any workbook produced by a C# application.
// AI Prompts: Write C# code with Aspose.Cells to add a Form button that triggers a VBA macro named HighlightActiveRow and saves the workbook as .xlsm. | Explain how to change the highlight color in the embedded VBA macro and handle potential runtime errors. | Show how to add multiple buttons, each linked to different macros for row and column highlighting, using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

// Creates a new workbook, inserts a Form button on the first worksheet, assigns the macro name HighlightActiveRow, embeds a VBA module that colors the active cell's entire row yellow, and saves the file as a macro‑enabled .xlsm using Aspose.Cells for .NET.
class HighlightActiveRowButtonDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a Form button to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
        Button button = sheet.Shapes.AddButton(2, 0, 2, 0, 30, 100);
        button.Text = "Highlight Row";
        button.MacroName = "HighlightActiveRow";

        // Ensure the workbook has a VBA project (required for macro-enabled files)
        // Adding a dummy module creates the project automatically
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "Module1");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];

        // VBA code that highlights the entire row of the active cell
        string vbaCode = @"
Sub HighlightActiveRow()
    On Error Resume Next
    ActiveCell.EntireRow.Interior.ColorIndex = 6   'Yellow
End Sub
";
        module.Codes = vbaCode;

        // Save as a macro‑enabled workbook
        workbook.Save("HighlightActiveRowButton.xlsm", SaveFormat.Xlsm);
    }
}
