// Title: Add a TextBox shape with custom text and blue font to the first worksheet using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, insert a TextBox at row 2 column 2 on the first sheet, set its text to "Hello, Aspose.Cells!", apply a 12‑point blue font, and save as Output.xlsx using Aspose.Cells in C#. | Use Aspose.Cells to add a TextBox shape to a worksheet, configure the TextBox.Text property, change Font.Size to 12 and Font.Color to blue, then export the workbook. | Generate an Excel file with a TextBox shape positioned at row 2, column 2, containing custom text and styled font, leveraging the Aspose.Cells.Drawing API in C#.
// Common Searches: how to insert a textbox shape in Aspose.Cells C# and set its text | Aspose.Cells set font size and color for textbox shape | example of adding a textbox to the first worksheet with Aspose.Cells for .NET | Aspose.Cells C# create textbox at specific row and column | save workbook with textbox shape using Aspose.Cells
// Tags: textbox shape insertion Aspose.Cells | configure textbox text Aspose.Cells | apply font styling to textbox Aspose.Cells | worksheet shape handling Aspose.Cells | export workbook with shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// // This program creates a new workbook, adds a TextBox shape to the first worksheet at row 2 column 2, sets its text to "Hello, Aspose.Cells!", applies a 12‑point blue font, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a TextBox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape textbox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

        // Set the initial text of the TextBox
        textbox.Text = "Hello, Aspose.Cells!";

        // Optional: set font properties
        textbox.Font.Size = 12;
        textbox.Font.Color = Color.Blue;

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
