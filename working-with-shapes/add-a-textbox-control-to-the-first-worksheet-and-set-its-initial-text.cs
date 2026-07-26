// Title: Add a TextBox to the first worksheet and set its text with Aspose.Cells for .NET
// Description: Creates a new Workbook, accesses the first Worksheet, inserts a TextBox shape at row 1, column 1 (50 px × 150 px), assigns the text "Hello, Aspose.Cells!", and saves the file as TextBoxDemo.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells TextBox | add TextBox shape .NET | set TextBox text Aspose.Cells | Aspose.Cells worksheet shapes | C# Aspose.Cells save workbook
// Common Searches: Aspose.Cells add TextBox to worksheet | How to set initial text of a TextBox in Aspose.Cells | C# create TextBox shape with Aspose.Cells | Save workbook after adding TextBox Aspose.Cells
// Developer Intent: Insert a TextBox shape into the first worksheet and define its initial content.
// Use Cases: Add a labeled box to a generated report for instructions or titles. | Create a placeholder comment area that will be filled with dynamic data later. | Provide a user‑editable field in a template before the workbook is distributed.
// AI Prompts: Write C# code with Aspose.Cells to place a TextBox at cell B2, size 60 × 200 px, and set its text from a variable. | Show how to load an existing workbook and modify the Text property of a specific TextBox. | Explain how to change font style, size, and background color of a TextBox shape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, accesses the first Worksheet, inserts a TextBox shape at row 1, column 1 (50 px × 150 px), assigns the text "Hello, Aspose.Cells!", and saves the file as TextBoxDemo.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a TextBox to the worksheet (row 1, column 1, height 50px, width 150px)
        int textBoxIndex = worksheet.TextBoxes.Add(1, 1, 50, 150);

        // Retrieve the added TextBox
        TextBox textBox = worksheet.TextBoxes[textBoxIndex];

        // Set the initial text of the TextBox
        textBox.Text = "Hello, Aspose.Cells!";

        // Save the workbook
        workbook.Save("TextBoxDemo.xlsx");
    }
}
