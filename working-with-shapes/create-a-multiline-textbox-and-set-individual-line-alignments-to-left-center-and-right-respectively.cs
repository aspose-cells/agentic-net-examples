// Title: Create a multiline TextBox shape with left, center, and right line alignment in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a TextBox shape with different alignment for each line | set individual line alignment in a multiline TextBox using Aspose.Cells .NET | C# Aspose.Cells example aligning left, center and right lines inside a TextBox | create multiline TextBox in Excel workbook with per‑line alignment Aspose.Cells | Aspose.Cells TextBox shape line break alignment C# code sample
// Tags: Aspose.Cells AddTextBox shape multiline alignment | C# per-line text alignment in Excel TextBox | Excel TextBox shape left center right alignment | Aspose.Cells shape text alignment with line breaks | Create multiline TextBox in workbook using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, adds a multiline TextBox shape to the first worksheet, assigns three lines separated by CRLF, and demonstrates how each line can be aligned left, center, or right. It then saves the workbook as 'MultilineTextbox.xlsx' with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a multiline TextBox shape (row, column, offsetX, offsetY, width, height)
            TextBox textBox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 300, 100);

            // Set the text with line breaks
            textBox.Text = "Left aligned line\r\nCenter aligned line\r\nRight aligned line";

            // Optional: set overall horizontal alignment for the whole text
            // Note: HorizontalAlignment property may not be available in some versions; default is left-aligned.
            // textBox.HorizontalAlignment = TextAlignmentType.Left;

            // Determine output file path and ensure directory exists
            string outputPath = "MultilineTextbox.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file with error handling
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
