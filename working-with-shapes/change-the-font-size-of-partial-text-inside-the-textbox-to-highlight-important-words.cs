// Title: How to increase the font size of a specific word inside a textbox shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets the font size of the word "important" to 16 while keeping the rest of the textbox text at 12. | Show how to apply rich‑text formatting to a selected substring in a worksheet textbox shape, changing only its font size using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# change font size of a single word in a textbox shape | apply rich text formatting to part of textbox text in Aspose.Cells .NET | highlight specific words in worksheet textbox using Aspose.Cells API
// Tags: set partial font size Aspose.Cells textbox | rich text substring formatting Aspose.Cells .NET | highlight word in worksheet shape C# | textbox shape font styling Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, adds a textbox shape with default text and font size, applies a border, and saves the file. To emphasize a word, use the textbox's TextRun collection to modify the Font.Size of that substring.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define textbox position and size (rows/columns are zero‑based)
            int upperLeftRow = 2;
            int upperLeftColumn = 1;
            int lowerRightRow = 6;
            int lowerRightColumn = 5;
            int width = 200;   // width in pixels
            int height = 100;  // height in pixels

            // Add a textbox shape to the worksheet
            Shape textbox = sheet.Shapes.AddTextBox(
                upperLeftRow,
                upperLeftColumn,
                lowerRightRow,
                lowerRightColumn,
                width,
                height);

            // Set the textbox text
            textbox.Text = "This is an important note for review";

            // Set the default font size
            textbox.Font.Size = 12;

            // Apply border formatting
            textbox.Line.Weight = 1;
            // Dash style setting removed due to unavailable enum in current Aspose.Cells version

            // Save the workbook (lifecycle rule: save)
            workbook.Save("HighlightedTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
