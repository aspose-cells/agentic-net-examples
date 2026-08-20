// Title: Insert a Form Control button into a specific worksheet cell using Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, selects the first sheet, and adds a Form Control button anchored to cell B2 via Shapes.AddButton, specifying row/column indices and pixel offsets. It sets the button's caption and name, optionally links a macro, and saves the result as FormButton.xlsx.
// Keywords: Aspose.Cells | C# AddButton | Form Control button | Excel button macro | anchor button to cell | Shapes.AddButton | programmatic Excel UI | worksheet button placement
// Common Searches: Aspose.Cells add button to cell | C# place Form Control button in Excel | assign macro to Aspose.Cells button | Shapes.AddButton method parameters | create interactive Excel button with Aspose.Cells
// Developer Intent: Programmatically embed a Form Control button at a defined cell location in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Add a clickable button to a generated report that launches a predefined macro when the file is opened. | Position a UI element on a dashboard sheet to trigger a data refresh routine. | Build automated Excel forms where each button executes custom downstream logic.
// AI Prompts: Generate C# code to place a Form Control button in cell C5 and bind it to a macro named RefreshData using Aspose.Cells. | Explain each parameter of Shapes.AddButton and how to calculate pixel offsets for exact placement. | Create a script that adds multiple buttons across a range, each with a unique name and macro, with Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, selects the first sheet, and adds a Form Control button anchored to cell B2 via Shapes.AddButton, specifying row/column indices and pixel offsets. It sets the button's caption and name, optionally links a macro, and saves the result as FormButton.xlsx.
class InsertFormButton
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a Form Control button anchored to cell B2 (row index 1, column index 1)
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);
        button.Text = "Press Me";
        button.Name = "MyButton";
        // Optional: assign a macro to the button
        // button.MacroName = "MyMacro";

        // Save the workbook
        workbook.Save("FormButton.xlsx");
    }
}
