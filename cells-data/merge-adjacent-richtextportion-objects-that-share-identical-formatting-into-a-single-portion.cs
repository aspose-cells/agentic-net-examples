// Title: C# – Merge Adjacent RichTextPortion Objects with Identical Formatting in Aspose.Cells TextBox
// Description: Demonstrates how to create a workbook, add a TextBox shape, apply the same font style to consecutive character ranges, detect identical Font properties, consolidate those ranges into a single RichTextPortion, and save the result. Reduces internal rich‑text objects and improves rendering performance.
// Keywords: Aspose.Cells RichTextPortion merge | C# Aspose.Cells TextBox formatting | combine consecutive font runs | reduce RichTextPortion count | Aspose.Cells merge identical font styles | TextBox rich text consolidation | Aspose.Cells .NET example
// Common Searches: merge adjacent RichTextPortion Aspose.Cells .NET | combine consecutive characters with same font in TextBox | reduce number of RichTextPortion objects in Excel workbook | Aspose.Cells C# merge text formatting runs | how to consolidate rich text portions in a shape
// Developer Intent: Identify consecutive characters that share the same Font settings and collapse them into a single RichTextPortion to simplify the text model and boost performance.
// Use Cases: After applying bold formatting to two neighboring ranges, merge them into one RichTextPortion to keep the document lightweight. | When generating reports with mixed styles, automatically combine runs that have identical Font attributes to streamline rendering. | Before exporting a workbook, re‑apply merged formatting so the TextBox contains the minimal number of RichTextPortion objects.
// AI Prompts: Write a C# method for Aspose.Cells that scans a Shape's Characters collection, compares Font properties, and merges adjacent RichTextPortion objects with identical formatting. | Generate code that extracts formatting runs from a TextBox, determines equality of Font attributes, and re‑applies merged runs to reduce RichTextPortion count. | Explain step‑by‑step how to compare two Aspose.Cells Font objects for equality and rebuild merged runs to simplify rich‑text structures in a workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsRichTextMergeDemo
{
    // Demonstrates how to create a workbook, add a TextBox shape, apply the same font style to consecutive character ranges, detect identical Font properties, consolidate those ranges into a single RichTextPortion, and save the result. Reduces internal rich‑text objects and improves rendering performance.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Add a TextBox shape with rich text ----------
                // Parameters: upper left row, upper left column, upper left offset, height, width, lower right offset
                Shape textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 100, 200, 0);
                // Sample text containing three logical portions
                textBox.Text = "BoldPart1BoldPart2NormalPart";

                // Apply formatting to create adjacent portions with identical formatting
                // Portion 1: characters 0-9 (Bold)
                textBox.Characters(0, 10).Font.IsBold = true;
                // Portion 2: characters 10-19 (Bold) – same formatting as previous, should be merged
                textBox.Characters(10, 10).Font.IsBold = true;
                // Portion 3: characters 20-30 (Normal)
                textBox.Characters(20, 11).Font.IsBold = false;

                // ---------- Merge adjacent RichTextPortion objects with identical formatting ----------
                string fullText = textBox.Text;
                int textLength = fullText.Length;

                // Helper method to compare two Font objects for equality of relevant properties
                bool FontsAreEqual(Font f1, Font f2)
                {
                    return f1.IsBold == f2.IsBold &&
                           f1.IsItalic == f2.IsItalic &&
                           f1.Underline == f2.Underline &&   // use Underline property
                           f1.Size == f2.Size &&
                           f1.Color.ToArgb() == f2.Color.ToArgb() &&
                           f1.Name == f2.Name;
                }

                // List to hold merged runs: start index, length, and the Font to apply
                var mergedRuns = new List<(int Start, int Length, Font Font)>();

                int index = 0;
                while (index < textLength)
                {
                    // Get the formatting of the current character
                    Font currentFont = textBox.Characters(index, 1).Font;

                    int runLength = 1;
                    // Extend the run while the next character has identical formatting
                    while (index + runLength < textLength)
                    {
                        Font nextFont = textBox.Characters(index + runLength, 1).Font;
                        if (!FontsAreEqual(currentFont, nextFont))
                            break;
                        runLength++;
                    }

                    // Store the merged run
                    mergedRuns.Add((index, runLength, currentFont));

                    // Move to the next unprocessed character
                    index += runLength;
                }

                // Re‑apply formatting based on the merged runs.
                // This effectively reduces the number of internal RichTextPortion objects.
                foreach (var run in mergedRuns)
                {
                    var chars = textBox.Characters(run.Start, run.Length);
                    chars.Font.IsBold = run.Font.IsBold;
                    chars.Font.IsItalic = run.Font.IsItalic;
                    chars.Font.Underline = run.Font.Underline;   // apply underline
                    chars.Font.Size = run.Font.Size;
                    chars.Font.Color = run.Font.Color;
                    chars.Font.Name = run.Font.Name;
                }

                // ---------- Save the workbook ----------
                workbook.Save("RichTextPortionMergeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
