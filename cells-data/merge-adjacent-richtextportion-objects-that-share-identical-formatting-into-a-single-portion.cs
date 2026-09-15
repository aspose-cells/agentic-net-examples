// Title: Consolidate adjacent RichTextPortion objects with identical Font formatting in a single Excel cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through a cell's Characters collection in Aspose.Cells, detects consecutive characters with the same Font properties, and merges them into one RichTextPortion. | Create a helper method that compares two Aspose.Cells Font objects for equality of name, size, style, underline, color, and strikeout. | Develop a console application that loads an XLSX workbook, consolidates redundant rich‑text formatting in cell A1, and saves the updated file while preserving the original text.
// Common Searches: Aspose.Cells C# merge consecutive rich text portions with same font in a cell | remove duplicate font formatting from Excel cell using Aspose.Cells .NET | how to combine adjacent RichTextPortion objects in Aspose.Cells workbook
// Tags: merge rich text portions Aspose.Cells | consolidate cell font formatting .NET | remove redundant rich text Aspose.Cells | compare Font objects Aspose.Cells | optimize rich text performance Excel .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRichTextMerge
{
    // The example loads an XLSX workbook, reads the text of cell A1, groups consecutive characters that share identical Font attributes, records each merged segment, rewrites the plain text to keep existing rich‑text, reapplies the consolidated font settings to each segment, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file existence
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];
                Cell cell = worksheet.Cells["A1"];

                // Retrieve plain text
                string text = cell.StringValue ?? string.Empty;
                if (text.Length == 0)
                {
                    Console.WriteLine("Cell A1 is empty.");
                    workbook.Save(outputPath);
                    return;
                }

                // Group consecutive characters with identical formatting
                List<(int Start, int Length, Font Font)> mergedPortions = new List<(int, int, Font)>();
                int startIdx = 0;
                Font prevFont = cell.Characters(0, 1).Font; // first character font

                for (int i = 1; i < text.Length; i++)
                {
                    Font curFont = cell.Characters(i, 1).Font;
                    if (!FontsEqual(prevFont, curFont))
                    {
                        mergedPortions.Add((startIdx, i - startIdx, prevFont));
                        startIdx = i;
                        prevFont = curFont;
                    }
                }

                // Add final portion
                mergedPortions.Add((startIdx, text.Length - startIdx, prevFont));

                // Write back plain text (preserves existing rich text)
                cell.PutValue(text);

                // Apply merged formatting to each portion
                foreach (var portion in mergedPortions)
                {
                    Font targetFont = cell.Characters(portion.Start, portion.Length).Font;
                    ApplyFontSetting(targetFont, portion.Font);
                }

                // Save workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Compare two Font objects for equality of key properties
        private static bool FontsEqual(Font f1, Font f2)
        {
            if (f1 == null || f2 == null) return false;

            return f1.Name == f2.Name &&
                   f1.Size == f2.Size &&
                   f1.IsBold == f2.IsBold &&
                   f1.IsItalic == f2.IsItalic &&
                   f1.Underline == f2.Underline &&
                   f1.Color.ToArgb() == f2.Color.ToArgb() &&
                   f1.IsStrikeout == f2.IsStrikeout;
        }

        // Copy relevant font attributes from source Font to target Font
        private static void ApplyFontSetting(Font target, Font source)
        {
            if (target == null || source == null) return;

            target.Name = source.Name;
            target.Size = source.Size;
            target.IsBold = source.IsBold;
            target.IsItalic = source.IsItalic;
            target.Underline = source.Underline;
            target.Color = source.Color;
            target.IsStrikeout = source.IsStrikeout;
        }
    }
}
