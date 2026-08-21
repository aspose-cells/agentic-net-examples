// Title: Add a Button Shape with Macro to Cell D5 using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a new workbook with Aspose.Cells, insert a Button shape anchored to cell D5, set its display text, assign a VBA macro name, and save the file as ButtonWithMacro.xlsx.
// Keywords: Aspose.Cells | C# button shape | Excel button macro | add button to cell D5 | assign macro to button | shape positioning Aspose.Cells | .NET Excel automation | VBA macro link | Workbook button control
// Common Searches: Aspose.Cells add button to specific cell | C# insert button shape at D5 using Aspose.Cells | Assign a macro name to an Excel button with Aspose.Cells | Set button text and size in Aspose.Cells .NET | How to place a macro‑linked button in a workbook programmatically
// Developer Intent: Insert a button control, link it to a macro, and position it on cell D5.
// Use Cases: Create a template workbook that includes a "Run Macro" button for end‑users to trigger data refresh. | Automate the addition of macro‑enabled buttons across multiple sheets in a reporting suite. | Build an interactive Excel dashboard where a button launches a predefined VBA routine.
// AI Prompts: Generate C# code with Aspose.Cells to add a button at D5, set its caption to "Run Macro", and assign the macro name "MyMacro". | Explain how to adjust height, width, and offset parameters when adding a button shape via Aspose.Cells. | Show the steps to embed a VBA macro in a workbook and bind it to a button created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example demonstrates how to create a new workbook with Aspose.Cells, insert a Button shape anchored to cell D5, set its display text, assign a VBA macro name, and save the file as ButtonWithMacro.xlsx.
class InsertButtonWithMacro
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the target cell (D5). Zero‑based indices: column D = 3, row 5 = 4
        int targetRow = 4;      // D5 row index
        int targetColumn = 3;   // D5 column index

        // Add a button anchored to the target cell.
        // top and left offsets are set to 0 pixels.
        // Height = 30 pixels, Width = 100 pixels (adjust as needed).
        Button button = sheet.Shapes.AddButton(
            topRow: targetRow,
            top: 0,
            leftColumn: targetColumn,
            left: 0,
            height: 30,
            width: 100);

        // Set button properties
        button.Text = "Run Macro";
        button.MacroName = "MyMacro";   // Assign the macro name

        // Save the workbook
        workbook.Save("ButtonWithMacro.xlsx");
    }
}
