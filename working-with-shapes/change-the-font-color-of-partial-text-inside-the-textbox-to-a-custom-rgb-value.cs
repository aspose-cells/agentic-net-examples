// Title: Apply a Custom RGB Font Color to Specific Characters in a TextBox Shape – Aspose.Cells for .NET
// Description: Creates a workbook, adds a textbox shape, defines a zero‑based character range, sets a Font with a custom RGB value, enables only the FontColor flag, and uses FormatCharacters to color that substring before saving the file.
// Keywords: Aspose.Cells partial text color | textbox shape format characters .NET | custom RGB font color Aspose.Cells | StyleFlag FontColor example | Excel shape text styling
// Common Searches: Aspose.Cells change color of part of textbox text | format specific characters in a shape using .NET | set RGB font color for substring in Excel textbox | partial text formatting Aspose.Cells
// Developer Intent: Color a selected range of characters inside a textbox shape with a custom RGB value.
// Use Cases: Highlight key terms in a generated report textbox. | Show status indicators like "Pass" or "Fail" in different colors within a shape. | Create a legend where particular words are uniquely colored for clarity.
// AI Prompts: Generate code to apply bold, italic, and a custom RGB color to a character range in an Aspose.Cells textbox. | Explain how to retrieve a textbox's existing Font, modify its color, and apply it to selected characters using StyleFlag. | Show how to change the background color of specific characters in a shape with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPartialTextColor
{
    // Creates a workbook, adds a textbox shape, defines a zero‑based character range, sets a Font with a custom RGB value, enables only the FontColor flag, and uses FormatCharacters to color that substring before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, upper left offset,
            // width, height (all in points)
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 50);
            textBox.Text = "Partial color change example";

            // Define the range of characters to format (e.g., characters 8 to 13 -> "color")
            int startIndex = 8;   // zero‑based index
            int length = 5;       // number of characters to format

            // Prepare a Font object with the desired custom RGB color
            // Here we use a teal color (R=0, G=128, B=128)
            Font font = textBox.Font;
            font.Color = Color.FromArgb(0, 128, 128);

            // Create a StyleFlag indicating that only the font color should be applied
            StyleFlag flag = new StyleFlag();
            flag.FontColor = true;

            // Apply the formatting to the specified character range
            textBox.FormatCharacters(startIndex, length, font, flag);

            // Save the workbook to a file
            workbook.Save("PartialTextColor.xlsx");
        }
    }
}
