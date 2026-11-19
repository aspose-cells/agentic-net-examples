// Title: Increase the font size of a specific word inside an Excel textbox shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change font size of a single word in a worksheet textbox | How to highlight a word in an Excel textbox using Aspose.Cells Characters method | Set partial text formatting in a textbox shape with Aspose.Cells for .NET | C# code to enlarge font of a substring inside an Aspose.Cells textbox | Apply different font size to part of text in Excel shape using Aspose.Cells
// Tags: Aspose.Cells textbox character range API | partial font size change in Excel shape C# | substring font styling with Aspose.Cells | C# manipulate textbox text formatting Aspose.Cells | Excel worksheet textbox font adjustment .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a textbox shape to the first worksheet, assigns full text, finds the target word, retrieves its character range via TextBox.Characters(start, length), sets Font.Size to highlight the word, and saves the workbook as HighlightedTextbox.xlsx.
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

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperLeftRow = 2;
            int upperLeftColumn = 1;
            int top = 5;      // pixels
            int left = 5;     // pixels
            int height = 80;  // pixels
            int width = 300;  // pixels

            TextBox textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);

            // Set the full text of the textbox
            string fullText = "Please review the important sections before submission.";
            textBox.Text = fullText;

            // Define the word to highlight
            string wordToHighlight = "important";

            // Find the start index of the word within the textbox text
            int startIndex = fullText.IndexOf(wordToHighlight, StringComparison.Ordinal);
            if (startIndex >= 0)
            {
                // Retrieve the characters range and apply a larger font size
                var chars = textBox.Characters(startIndex, wordToHighlight.Length);
                chars.Font.Size = 14; // Highlight by increasing font size
                // Optional styling:
                // chars.Font.Color = System.Drawing.Color.Red;
                // chars.Font.IsBold = true;
            }

            // Define output file path
            string outputPath = "HighlightedTextbox.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
