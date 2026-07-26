// Title: Add a Button Shape to Cell D5 with Macro Assignment using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a Button control at cell D5 (row 5, column D) with a 30 × 100 px size, sets its caption to "Run Macro", assigns the macro "MyMacro", and saves the file as ButtonWithMacro.xlsx.
// Keywords: Aspose.Cells add button | C# button shape Aspose.Cells | macro button Aspose.Cells | position button D5 | Aspose.Cells Drawing API
// Common Searches: Aspose.Cells add button to specific cell | Assign macro to button shape C# | Set button size and location Aspose.Cells | Create clickable macro button in Excel with Aspose
// Developer Intent: Insert a button at D5, define its text and macro, then save the workbook.
// Use Cases: Generate a template workbook that includes a ready‑to‑run macro button. | Automate report files that need a user‑triggered VBA action via a button. | Build interactive dashboards where each cell‑aligned button launches custom logic.
// AI Prompts: Write C# code with Aspose.Cells to place a button at D5, set its caption to "Run Macro", assign macro "MyMacro", and save the workbook. | Explain how to modify the height, width, and pixel offsets of a button added with Aspose.Cells. | Show how to locate an existing button shape in a worksheet and change its MacroName property programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a Button control at cell D5 (row 5, column D) with a 30 × 100 px size, sets its caption to "Run Macro", assigns the macro "MyMacro", and saves the file as ButtonWithMacro.xlsx.
class InsertButtonWithMacro
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a Button control positioned at cell D5 (row index 4, column index 3)
        // topRow = 4 (row 5), leftColumn = 3 (column D), offsets = 0, size = 30x100 pixels
        Button button = sheet.Shapes.AddButton(
            topRow: 4,    // Upper left row index (zero‑based)
            top: 0,       // Vertical offset in pixels
            leftColumn: 3,// Upper left column index (zero‑based)
            left: 0,      // Horizontal offset in pixels
            height: 30,   // Height in pixels
            width: 100);  // Width in pixels

        // Set the button's display text
        button.Text = "Run Macro";

        // Assign the macro name that will be executed when the button is clicked
        button.MacroName = "MyMacro";

        // Save the workbook
        workbook.Save("ButtonWithMacro.xlsx");
    }
}
