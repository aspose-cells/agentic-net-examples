// Title: Aspose.Cells .NET: Create a multiline TextBox with left, center, and right alignment per line
// Description: Demonstrates how to add a TextBox shape to a worksheet, insert three newline‑separated lines, access the TextParagraphCollection, and set the AlignmentType of each paragraph to Left, Center, and Right before saving the workbook.
// Keywords: Aspose.Cells multiline TextBox | C# TextBox paragraph alignment | Aspose.Cells TextAlignmentType | TextParagraphCollection example | .NET spreadsheet shape alignment | Aspose.Cells left center right alignment
// Common Searches: Aspose.Cells set different alignment for each line in a TextBox | C# multiline TextBox alignment Aspose.Cells | How to align paragraphs inside a TextBox using Aspose.Cells | Aspose.Cells TextBox left center right alignment code
// Developer Intent: Add a TextBox with three lines and apply distinct left, center, and right alignment to each line in a .NET workbook.
// Use Cases: Create a report header where the title is centered, a subtitle is left‑aligned, and a signature line is right‑aligned within a single TextBox. | Design a data‑entry form that shows instructions with varied alignment to guide users, using one TextBox for all paragraphs. | Generate an invoice where item descriptions are left‑aligned, totals are centered, and the authorized signatory line is right‑aligned inside one TextBox.
// AI Prompts: Show C# code that adds a multiline TextBox to an Aspose.Cells worksheet and sets the first line left‑aligned, the second line centered, and the third line right‑aligned. | Provide an Aspose.Cells example that accesses a TextBox's TextParagraphCollection and changes each paragraph's AlignmentType to Left, Center, and Right. | Explain how to create a three‑paragraph TextBox in Aspose.Cells for .NET and apply different alignments to each paragraph.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a TextBox shape to a worksheet, insert three newline‑separated lines, access the TextParagraphCollection, and set the AlignmentType of each paragraph to Left, Center, and Right before saving the workbook.
    public class MultilineTextBoxAlignmentDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper row, left column, height (pixels), width (pixels)
            int textboxIndex = worksheet.TextBoxes.Add(0, 0, 150, 300);
            TextBox textbox = worksheet.TextBoxes[textboxIndex];

            // Set multiline text (each line separated by newline)
            textbox.Text = "Left aligned line\nCenter aligned line\nRight aligned line";

            // Access the paragraphs (each line is a separate paragraph)
            TextParagraphCollection paragraphs = textbox.TextBody.TextParagraphs;

            // Set individual alignment for each paragraph
            if (paragraphs.Count >= 3)
            {
                paragraphs[0].AlignmentType = TextAlignmentType.Left;    // First line -> left
                paragraphs[1].AlignmentType = TextAlignmentType.Center; // Second line -> center
                paragraphs[2].AlignmentType = TextAlignmentType.Right;  // Third line -> right
            }

            // Ensure output directory exists
            string outputPath = "MultilineTextBoxAlignmentDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
