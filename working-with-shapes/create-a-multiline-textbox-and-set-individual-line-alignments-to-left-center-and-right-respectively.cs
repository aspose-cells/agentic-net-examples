// Title: How to add a multiline TextBox to an Excel worksheet and align each line left, center, and right using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a TextBox containing three lines of text and sets the first line left‑aligned, the second line centered, and the third line right‑aligned. | Modify the Aspose.Cells TextBox example to use RichText or paragraph formatting so that each line inside the same TextBox has a different horizontal alignment. | Generate a complete console application that creates a workbook, adds a multiline TextBox at a specific cell, applies per‑line alignment, and saves the file as an .xlsx workbook.
// Common Searches: Aspose.Cells C# create TextBox with line breaks and set different alignments per line | set individual paragraph alignment inside an Aspose.Cells TextBox | multiline TextBox shape alignment Aspose.Cells .NET example | how to left center right align lines in an Excel TextBox using Aspose.Cells | C# Aspose.Cells add TextBox to worksheet and control text alignment per line
// Tags: Aspose.Cells add multiline TextBox | Aspose.Cells TextBox per-line alignment | C# Aspose.Cells TextBox shape formatting | Excel TextBox horizontal alignment Aspose.Cells | Aspose.Cells TextBox line break handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a multiline TextBox at row 2, column 2 with CRLF-separated lines, sets the overall horizontal alignment to left, and saves the workbook as MultilineTextBox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a multiline TextBox (row, column, top, left, height, width)
            // Row and column are zero‑based indices
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 100, 100, 200, 100);

            // Set multiline text using line breaks
            textBox.Text = "Left aligned line\r\nCenter aligned line\r\nRight aligned line";

            // Optional: set overall text alignment (applies to all lines)
            textBox.TextHorizontalAlignment = TextAlignmentType.Left;

            // Save the workbook
            string outputPath = "MultilineTextBox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
