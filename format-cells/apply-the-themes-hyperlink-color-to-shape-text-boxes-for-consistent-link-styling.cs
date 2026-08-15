// Title: Apply Workbook Theme Hyperlink Color to TextBox Shape Text with Aspose.Cells for .NET
// Description: Demonstrates how to retrieve the workbook's theme hyperlink color using GetThemeColor, create a style with that color, and apply it to all characters of a TextBox shape via StyleFlag and FormatCharacters. The example adds a clickable TextBox and saves the workbook.
// Keywords: Aspose.Cells C# | TextBox shape hyperlink color | GetThemeColor Hyperlink | StyleFlag FontColor | FormatCharacters Aspose.Cells | apply theme color to shape text | Excel shape formatting .NET | hyperlink styling in workbook | theme based text color Aspose | programmatic Excel hyperlink theme
// Common Searches: Aspose.Cells set hyperlink color for TextBox shape | How to use workbook theme hyperlink color in C# | Format TextBox characters with theme color Aspose.Cells | GetThemeColor Hyperlink example .NET | Apply StyleFlag to shape text in Aspose.Cells
// Developer Intent: The developer wants to color the text inside a TextBox shape using the workbook’s theme hyperlink color for consistent link styling.
// Use Cases: Generate reports where TextBox links automatically match the workbook’s hyperlink theme. | Create interactive dashboards with clickable TextBox shapes that follow the document’s color scheme. | Batch‑apply theme‑based hyperlink colors to multiple shape texts across a workbook.
// AI Prompts: Write C# code that loops through all TextBox shapes in a workbook and applies the theme hyperlink color to their text using Aspose.Cells. | Show how to retrieve the visited‑hyperlink theme color and use it for shape text styling in Aspose.Cells for .NET. | Explain the use of StyleFlag with FormatCharacters to change only the font color of a TextBox while preserving other formatting.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to retrieve the workbook's theme hyperlink color using GetThemeColor, create a style with that color, and apply it to all characters of a TextBox shape via StyleFlag and FormatCharacters. The example adds a clickable TextBox and saves the workbook.
class ApplyThemeHyperlinkColorToTextBox
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 50, 100, 200, 60);

            // Set sample text in the TextBox
            textBox.Text = "Visit Aspose";

            // Add a hyperlink to the entire shape (optional)
            textBox.AddHyperlink("https://www.aspose.com");

            // Retrieve the theme's hyperlink color
            Color themeHyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

            // Create a Font object with the desired color
            Style style = workbook.CreateStyle();
            Font hyperlinkFont = style.Font;
            hyperlinkFont.Color = themeHyperlinkColor;

            // Define a StyleFlag indicating that only the font color should be changed
            StyleFlag flag = new StyleFlag
            {
                FontColor = true
            };

            // Apply the font color to all characters in the TextBox
            int textLength = textBox.Text.Length;
            textBox.FormatCharacters(0, textLength, hyperlinkFont, flag);

            // Save the workbook
            workbook.Save("HyperlinkStyledTextBox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
