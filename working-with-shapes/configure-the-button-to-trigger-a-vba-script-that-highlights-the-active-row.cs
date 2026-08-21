// Title: Add a Worksheet Button that Runs a VBA Macro to Highlight the Active Row (Aspose.Cells for .NET)
// Description: Shows how to create a macro‑enabled XLSM workbook with Aspose.Cells, add a button shape, embed a procedural VBA module, link the macro "HighlightActiveRow" to the button, and shade the active row yellow on click.
// Keywords: Aspose.Cells add button VBA | C# macro enabled workbook | assign VBA macro to shape Aspose.Cells | highlight active row Excel | button click VBA highlight row | Aspose.Cells VBA integration | create XLSM with button C#
// Common Searches: Aspose.Cells assign VBA macro to button | C# create macro enabled Excel file with button | how to highlight active row with VBA button using Aspose.Cells | add procedural VBA module in Aspose.Cells C# | button shape run macro Aspose.Cells .NET
// Developer Intent: Generate an XLSM file that contains a button which executes a VBA macro to color the active row.
// Use Cases: One‑click row highlighting for reviewers of auto‑generated reports. | Interactive data‑entry sheets where users can quickly mark the row they are editing. | Template designs that let end‑users emphasize rows without manual formatting.
// AI Prompts: Write C# code with Aspose.Cells to insert a button that runs a VBA macro highlighting the active row. | Explain how to embed a procedural VBA module and bind its macro to a shape in a macro‑enabled workbook using Aspose.Cells. | Show how to modify the VBA macro to use a different fill color and update the button caption accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsButtonVbaDemo
{
    // Shows how to create a macro‑enabled XLSM workbook with Aspose.Cells, add a button shape, embed a procedural VBA module, link the macro "HighlightActiveRow" to the button, and shade the active row yellow on click.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button to the worksheet (row 1, column 1, size 100x30 pixels)
            Button button = sheet.Shapes.AddButton(0, 0, 0, 0, 100, 30);
            button.Text = "Highlight Row";
            button.Name = "HighlightButton";

            // Set the macro name that will be executed when the button is clicked
            button.MacroName = "HighlightActiveRow";

            // Ensure the workbook has a VBA project (required for macro-enabled files)
            // Adding a procedural module to hold the macro code
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "Module1");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // VBA code: highlights the entire active row with yellow background
            string vbaCode = @"
Sub HighlightActiveRow()
    ActiveCell.EntireRow.Interior.Color = vbYellow
End Sub
";
            module.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file
            workbook.Save("ButtonWithVba.xlsm", SaveFormat.Xlsm);
        }
    }
}
