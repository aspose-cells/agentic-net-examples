// Title: Apply Custom RGB Font Color to Part of Text in a TextBox Shape with Aspose.Cells for .NET
// Description: Shows how to color a specific word inside a textbox shape using a custom RGB value via Style, StyleFlag, and TextBody.Format in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | textbox shape | partial text color | custom RGB | StyleFlag | TextBody.Format | font color substring | Excel shape formatting | generated report styling
// Common Searches: Aspose.Cells change color of part of textbox text | partial text formatting in Excel shape .NET | custom RGB font color textbox Aspose.Cells | how to color a word in a shape using Aspose.Cells | TextBody.Format example Aspose.Cells
// Developer Intent: Set a custom RGB font color for a selected substring inside a textbox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Highlight a keyword in a generated report to draw reader attention. | Display a brand name in a distinct color within a marketing worksheet. | Differentiate label sections by applying unique RGB colors to each part.
// AI Prompts: Provide C# code that changes the word "Aspose" in a textbox to RGB #800080 using Aspose.Cells. | Show how to apply multiple custom RGB colors to different substrings within a textbox shape with Aspose.Cells for .NET. | Explain the role of StyleFlag and TextBody.Format when formatting partial text in a shape.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to color a specific word inside a textbox shape using a custom RGB value via Style, StyleFlag, and TextBody.Format in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 300, 100);
        textBox.Text = "Hello Aspose.Cells!";

        // Define a custom RGB color (e.g., purple: 128, 0, 128)
        Color customRgb = Color.FromArgb(128, 0, 128);

        // Create a style and set its font color to the custom RGB value
        Style style = workbook.CreateStyle();
        style.Font.Color = customRgb;

        // Create a StyleFlag indicating that only the font color should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontColor = true;

        // Determine the start index and length of the text segment to color
        // Here we color the word "Aspose"
        int startIndex = textBox.Text.IndexOf("Aspose");
        int length = "Aspose".Length;

        // Apply the formatting to the selected characters using the TextBody.Format method
        textBox.TextBody.Format(startIndex, length, style.Font, flag);

        // Save the workbook (output file will contain the textbox with partially colored text)
        workbook.Save("TextboxPartialColor.xlsx");
    }
}
