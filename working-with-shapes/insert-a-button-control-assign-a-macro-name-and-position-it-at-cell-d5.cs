// Title: Add a macro‑linked button control to cell D5 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook with Aspose.Cells, inserts a button shape at cell D5, sets its MacroName to "MyMacro", assigns the caption "Click Me", and saves the file as Output.xlsx. | Provide a .NET example that uses Aspose.Cells to place a form‑control button on D5, bind it to a macro, and customize its displayed text.
// Common Searches: Aspose.Cells C# how to place a form control button in a specific cell | assign a macro to a button created with Aspose.Cells in a .NET project | C# code to add a button at D5 and set its caption using Aspose.Cells library | save workbook after inserting macro‑enabled button with Aspose.Cells
// Tags: add button shape Aspose.Cells C# | macro name assignment Aspose.Cells button | button placement at cell D5 Aspose.Cells | custom button caption Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, adds a button anchored to cell D5, sets MacroName="MyMacro" and Text="Click Me", then saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell position for the button (D5)
        int row = 4;      // Zero‑based row index for D5
        int column = 3;   // Zero‑based column index for D5

        // Define button size (in pixels)
        int height = 30;
        int width = 80;

        // Add a button shape anchored to cell D5
        Button button = sheet.Shapes.AddButton(row, column, 0, 0, height, width);

        // Assign a macro name to the button
        button.MacroName = "MyMacro";

        // Optional: set the button caption
        button.Text = "Click Me";

        // Save the workbook
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
