// Title: Compress text in a narrow TextBox shape by applying negative character spacing with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set negative character spacing for TextBox shape C# | How to compress long text inside an Excel TextBox using Aspose.Cells | C# example of TextBox.TextOptions.Spacing property in Aspose.Cells | Reduce character spacing in an Excel shape with Aspose.Cells .NET | Apply custom spacing to TextBox text in an Aspose.Cells workbook
// Tags: TextOptions.Spacing property Aspose.Cells | Excel shape text compression .NET | narrow TextBox formatting Aspose.Cells | adjust shape font spacing C# | apply custom spacing to TextBox shape

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ApplyCharacterSpacing
{
    // The sample creates a workbook, adds a narrow TextBox shape, assigns a long word, applies a -2 point character spacing via TextBox.TextOptions.Spacing to compress the text, sets font styling, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "CompressedTextShape.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a narrow TextBox shape (height 30 pixels, width 80 pixels)
                // Parameters: top row, left column, height (pixels), width (pixels)
                int textboxIndex = sheet.TextBoxes.Add(2, 1, 30, 80);
                TextBox textbox = sheet.TextBoxes[textboxIndex];

                // Set the text that will be compressed
                textbox.Text = "VeryLongWordWithoutSpaces";

                // Apply custom character spacing (negative value compresses text)
                // TextOptions controls spacing for shapes
                textbox.TextOptions.Spacing = -2.0;

                // Optional: set other font attributes for better visibility
                textbox.Font.Name = "Arial";
                textbox.Font.Size = 12;
                textbox.Font.IsBold = true;
                textbox.Font.Color = Color.DarkBlue;

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
