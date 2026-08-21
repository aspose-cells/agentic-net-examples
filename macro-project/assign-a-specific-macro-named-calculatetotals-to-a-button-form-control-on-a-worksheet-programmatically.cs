// Title: Assign the CalculateTotals macro to a worksheet button using Aspose.Cells C#
// Description: Creates a new workbook, adds a button shape to the first worksheet, sets its caption, links it to the VBA macro "CalculateTotals" via the MacroName property, enables macros, and saves the file as an XLSM workbook.
// Keywords: Aspose.Cells C# | macro enabled workbook | button shape | MacroName property | assign VBA macro to button | Excel automation | XLSM file | programmatic UI control | CalculateTotals macro
// Common Searches: Aspose.Cells assign VBA macro to button C# | How to set button.MacroName in Aspose.Cells | Create XLSM workbook with button that runs macro | Link a macro to a shape using Aspose.Cells .NET | Programmatically add a button that calls CalculateTotals
// Developer Intent: Link the VBA macro "CalculateTotals" to a button control on a worksheet and save the result as a macro‑enabled Excel file.
// Use Cases: Add a "Calculate" button that triggers a total‑calculation macro in generated reports. | Build interactive Excel templates where users launch predefined VBA procedures with a click. | Automate the creation of macro‑enabled workbooks that include UI controls linked to specific macros.
// AI Prompts: Show C# code with Aspose.Cells that adds a button to a sheet and assigns the macro "CalculateTotals". | Generate an example of creating an XLSM workbook, inserting a button, setting its MacroName, and enabling macros. | Explain how to modify an existing Aspose.Cells workbook to attach the "CalculateTotals" macro to an existing button shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a button shape to the first worksheet, sets its caption, links it to the VBA macro "CalculateTotals" via the MacroName property, enables macros, and saves the file as an XLSM workbook.
class AssignMacroToButton
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a button to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
        button.Text = "Calculate";

        // Assign the macro name to the button
        button.MacroName = "CalculateTotals";

        // Ensure macros are enabled in the workbook
        workbook.Settings.EnableMacros = true;

        // Save the workbook as a macro‑enabled file
        workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
    }
}
