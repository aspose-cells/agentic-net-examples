// Title: Add a TextBox Shape to a Worksheet Cell with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, locate a cell using zero‑based row and column indices, add a TextBox shape with defined pixel height and width, assign its text, and save the result as AddTextbox.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# textbox shape | add textbox Aspose.Cells | textbox at cell coordinates | worksheet shape positioning | Aspose.Cells TextBox API | Excel automation C# | zero based indices | pixel dimensions | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells add textbox to specific cell | C# place textbox at row and column in Excel | How to set textbox size with Aspose.Cells | Aspose.Cells TextBox example .NET | Add shape to worksheet using Aspose.Cells | GitHub Aspose.Cells textbox sample | Aspose.Cells tutorial for US developers
// Developer Intent: Insert a TextBox shape at a specified row‑column location in an Excel worksheet.
// Use Cases: Add an explanatory note next to a data cell in an automatically generated report. | Overlay a form‑field‑like box on a target cell for interactive Excel dashboards. | Highlight a key metric with a callout box in a financial summary workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a textbox at row 5, column 3, using custom height, width, and formatted text. | Provide an example that creates multiple textboxes in different cells and aligns them uniformly with Aspose.Cells. | Show how to change a textbox's font style, border, and background color after adding it using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, locate a cell using zero‑based row and column indices, add a TextBox shape with defined pixel height and width, assign its text, and save the result as AddTextbox.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the cell coordinates (zero‑based indices) where the textbox will be placed
        int topRow = 2;      // Row index (e.g., third row)
        int leftColumn = 1;  // Column index (e.g., second column)

        // Define the size of the textbox in pixels
        int height = 100;
        int width = 200;

        // Add a textbox to the worksheet using the TextBoxes collection
        int textboxIndex = worksheet.TextBoxes.Add(topRow, leftColumn, height, width);
        TextBox textbox = worksheet.TextBoxes[textboxIndex];
        textbox.Text = "Sample TextBox";

        // Save the workbook to a file
        workbook.Save("AddTextbox.xlsx");
    }
}
