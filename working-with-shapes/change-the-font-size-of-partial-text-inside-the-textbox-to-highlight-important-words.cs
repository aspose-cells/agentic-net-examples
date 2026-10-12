// Title: Increase the font size of a specific word inside an Aspose.Cells TextBox with C#
// AI Prompts: Write C# code that creates a TextBox on an Excel worksheet using Aspose.Cells and enlarges only the word "important" while leaving the rest of the text at the default size. | Show how to apply a larger font size to a substring of a TextBox's text in Aspose.Cells for .NET, including setting the text, locating the target word, and adjusting its Font.Size property. | Generate an example that demonstrates partial rich‑text formatting (font size change) for a word inside an Aspose.Cells TextBox shape.
// Common Searches: Aspose.Cells C# change font size of a single word in a textbox | how to format part of textbox text with larger font using Aspose.Cells .NET | C# Aspose.Cells set different font size for substring in Excel textbox | partial rich text styling in Aspose.Cells TextBox example | increase font size of specific word inside Excel textbox Aspose.Cells
// Tags: textbox partial font size Aspose.Cells | set substring font size C# Aspose.Cells | highlight word in Excel textbox .NET | rich text formatting Aspose.Cells TextBox | partial text styling Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a TextBox shape to the first worksheet, sets its text to "This is important text.", applies a uniform font size of 16 to the entire box, and saves the file as HighlightedTextBox.xlsx. It serves as a basis for applying richer formatting—such as increasing the font size of only the word "important"—by using Aspose.Cells' rich‑text capabilities on character ranges.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape (row, column, upper‑left row offset, upper‑left column offset, width, height)
            // In newer Aspose.Cells versions AddTextBox returns a TextBox directly.
            TextBox textBox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 300, 100);

            if (textBox == null)
                throw new InvalidOperationException("Failed to create TextBox shape.");

            // Set the textbox text
            textBox.Text = "This is important text.";

            // Apply formatting to the whole text (larger font size to emphasize)
            textBox.Font.Size = 16;

            // Define output path
            string outputPath = "HighlightedTextBox.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
