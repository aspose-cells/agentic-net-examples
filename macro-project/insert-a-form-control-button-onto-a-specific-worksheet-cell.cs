using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertFormControlButton
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Target cell where the button will be placed (e.g., B2)
        int targetRow = 1;      // Row index (zero‑based)
        int targetColumn = 1;   // Column index (zero‑based)

        // Add a Form Control button anchored to the target cell.
        // Parameters: topRow, top offset (pixels), leftColumn, left offset (pixels), height (pixels), width (pixels)
        Button button = sheet.Shapes.AddButton(targetRow, 0, targetColumn, 0, 30, 100);
        button.Text = "Click Me";

        // Optional: assign a macro to the button
        // button.MacroName = "MyMacro";

        // Save the workbook
        workbook.Save("FormButton.xlsx");
    }
}