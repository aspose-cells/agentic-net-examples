// Title: Create a TextBox shape on Sheet1 at cell B2 with a 200‑point width using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a TextBox shape to the first worksheet at cell B2, sets its width to 200 points, and saves the workbook. | Provide Aspose.Cells commands to position a TextBox at B2 on Sheet1 with custom width and height in points. | Write a C# snippet that creates a workbook, inserts a TextBox at B2, defines its size in points, adds sample text, and writes the file.
// Common Searches: Aspose.Cells C# add textbox to worksheet at specific cell | set textbox width in points using Aspose.Cells | how to position a shape at B2 cell with Aspose.Cells | C# Aspose.Cells shape size and location example | create textbox shape on Sheet1 with custom dimensions Aspose.Cells
// Tags: insert textbox shape Aspose.Cells | set textbox width points Aspose.Cells | place shape at cell B2 Aspose.Cells | textbox dimensions C# Aspose.Cells | worksheet shape creation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a workbook, inserts a TextBox shape on Sheet1 at cell B2 with a width of 200 points (height 100 points), sets sample text, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (named Sheet1 by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Define the position: cell B2 (row index 1, column index 1)
        int upperLeftRow = 1;      // B2 row (zero‑based)
        int upperLeftColumn = 1;   // B2 column (zero‑based)

        // Define size in points
        int height = 100; // height of the TextBox
        int width = 200;  // width of the TextBox as required

        // Add a TextBox shape at the specified cell with the given size
        Shape textbox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, 0, 0, height, width);

        // Optional: set some default text
        textbox.Text = "Sample TextBox";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
