// Title: Insert a TextBox at B2 with 200‑pt width using Aspose.Cells .NET
// Description: This C# example creates a new Workbook, accesses the first worksheet, adds a TextBox anchored to cell B2, sets its WidthPt property to 200 points, optionally assigns text, and saves the result as TextBoxAtB2.xlsx.
// Keywords: Aspose.Cells | C# TextBox | Insert TextBox B2 | WidthPt 200 | shape positioning .NET | Excel TextBox Aspose | worksheet shapes | Aspose.Cells example
// Common Searches: Aspose.Cells add TextBox to specific cell | set TextBox width in points C# | position shape at B2 Aspose.Cells | change TextBox WidthPt property | C# Aspose.Cells create textbox
// Developer Intent: Add a TextBox to Sheet1 at cell B2 and set its width to 200 points using Aspose.Cells.
// Use Cases: Create a header label in a financial report by placing a 200‑pt TextBox at B2. | Provide on‑sheet instructions for data entry with a fixed‑width TextBox anchored at B2. | Design a dashboard annotation where the TextBox width must match column layout, positioned at B2.
// AI Prompts: Generate C# code with Aspose.Cells that adds a TextBox at cell C3, height 60 px, width 150 pt. | Show how to change the background color and border style of a TextBox placed at D5 using Aspose.Cells. | Write a method that receives a list of cell addresses and adds a TextBox to each, setting each WidthPt to 180 points.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a new Workbook, accesses the first worksheet, adds a TextBox anchored to cell B2, sets its WidthPt property to 200 points, optionally assigns text, and saves the result as TextBoxAtB2.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (Sheet1 is the first worksheet)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a TextBox positioned at cell B2 (row index 1, column index 1)
        // Parameters: topRow, leftColumn, height (pixels), width (pixels)
        int textBoxIndex = sheet.TextBoxes.Add(1, 1, 50, 100);

        // Retrieve the created TextBox object
        TextBox textBox = sheet.TextBoxes[textBoxIndex];

        // Set the width of the TextBox to 200 points
        textBox.WidthPt = 200;

        // (Optional) Set some sample text
        textBox.Text = "Sample TextBox";

        // Save the workbook to a file
        workbook.Save("TextBoxAtB2.xlsx");
    }
}
