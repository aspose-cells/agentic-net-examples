// Title: Add a TextBox to the first worksheet and set its text using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a new Workbook, access the first Worksheet, insert a TextBox shape at row 1, column 1 (50 px × 150 px), assign the text "Hello, Aspose!", and save the file as TextBoxDemo.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells TextBox C# | add TextBox shape Aspose.Cells | set TextBox text .NET | Aspose.Cells worksheet shapes | C# Excel TextBox example | Aspose.Cells API TextBox
// Common Searches: how to insert a TextBox in Excel using Aspose.Cells C# | Aspose.Cells set initial text for TextBox shape | C# add textbox to worksheet at specific cell Aspose | Aspose.Cells create TextBox with dimensions
// Developer Intent: Insert a TextBox control into the first worksheet of a workbook and define its default text programmatically.
// Use Cases: Add a labeled instruction box to a financial report template. | Pre‑populate a form field in an Excel dashboard with placeholder text. | Highlight a key metric by placing a TextBox with custom wording on a summary sheet.
// AI Prompts: Generate C# code that adds a multiline TextBox to a given cell range, sets font size, color, and alignment using Aspose.Cells. | Show how to bind a TextBox shape to a worksheet cell so the text updates automatically when the cell value changes. | Provide an example that creates several TextBox shapes at different positions, each with unique initial text, in a single workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example demonstrates how to create a new Workbook, access the first Worksheet, insert a TextBox shape at row 1, column 1 (50 px × 150 px), assign the text "Hello, Aspose!", and save the file as TextBoxDemo.xlsx with Aspose.Cells for .NET.
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

        // Retrieve the added TextBox object
        TextBox textBox = worksheet.TextBoxes[textBoxIndex];

        // Set the initial text of the TextBox
        textBox.Text = "Hello, Aspose!";

        // Save the workbook
        workbook.Save("TextBoxDemo.xlsx");
    }
}
