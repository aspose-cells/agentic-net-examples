// Title: C# – Assign a VBA macro to a Forms button using Aspose.Cells
// Description: Demonstrates how to create a macro‑enabled workbook, add a VBA module with a simple Sub, insert a Forms button, set its MacroName property to the new macro, and save the file as .xlsm with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# macro button | set button MacroName Aspose.Cells | add VBA module programmatically | Forms button VBA macro .NET | macro‑enabled workbook Aspose | Excel automation C# Aspose.Cells | link VBA macro to shape | C# Excel button click macro
// Common Searches: How to bind a VBA macro to a Forms button with Aspose.Cells C# | Set MacroName property of a button in a macro‑enabled workbook | Add VBA module and assign its Sub to a shape using Aspose.Cells | Create .xlsm file with button that runs a macro in C# | Aspose.Cells example linking button to macro
// Developer Intent: The developer needs to programmatically connect a newly created VBA Sub to a Forms button so that clicking the button runs the macro in a macro‑enabled Excel file.
// Use Cases: Generate interactive Excel templates where users trigger custom VBA actions via on‑sheet buttons. | Automate report generators that include a “Refresh” button linked to a macro for data processing. | Build Excel‑based UI components (e.g., dashboards) that launch predefined VBA routines when clicked.
// AI Prompts: Write C# code with Aspose.Cells that adds a Forms button and assigns an existing VBA macro to its MacroName property. | Explain how to reference a macro located in a different VBA module when setting the MacroName of a button. | Provide steps to test that the button correctly executes the macro after saving the workbook as .xlsm.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

// Demonstrates how to create a macro‑enabled workbook, add a VBA module with a simple Sub, insert a Forms button, set its MacroName property to the new macro, and save the file as .xlsm with Aspose.Cells for .NET.
class SetButtonMacroDemo
{
    static void Main()
    {
        // Create a new workbook and enable macros
        Workbook workbook = new Workbook();
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a VBA module to the workbook and insert a simple macro
        int moduleIndex = workbook.VbaProject.Modules.Add(sheet);
        VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];
        vbaModule.Name = "Module1";
        vbaModule.Codes = 
            "Sub MyButtonMacro()\n" +
            "    MsgBox \"Button clicked!\"\n" +
            "End Sub";

        // Add a Forms button to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
        button.Text = "Run Macro";

        // Set the button's MacroName property to reference the newly added macro
        button.MacroName = "MyButtonMacro";

        // Save the workbook as a macro‑enabled file
        workbook.Save("ButtonWithMacro.xlsm");
    }
}
