// Title: C# – Merge Adjacent RichTextPortion Runs with Identical Formatting in Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a rectangle shape, adds rich‑text with alternating bold/normal segments, extracts FontSetting runs via GetRichFormattings, merges consecutive runs that share the same Font and TextOptions, reapplies the consolidated formatting, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | RichTextPortion | merge formatting runs | GetRichFormattings | FontSetting | shape rich text | consolidate text formatting | performance optimization | reduce file size
// Common Searches: Aspose.Cells merge RichTextPortion runs | C# combine adjacent rich text portions | How to consolidate identical font formatting in Aspose.Cells shape | Remove duplicate formatting from shape text Aspose.Cells | GetRichFormattings example C#
// Developer Intent: Combine consecutive RichTextPortion objects that have the same Font and TextOptions into a single run.
// Use Cases: Simplify shape text by reducing the number of formatting runs | Improve workbook load/save performance and file size after rich‑text editing | Prepare user‑generated shape text for export or further processing | Ensure consistent formatting when programmatically editing shapes
// AI Prompts: Write a reusable C# method that takes a Shape and merges its RichTextPortion runs with identical Font and TextOptions using Aspose.Cells. | Show how to iterate over FontSetting[] from GetRichFormattings, combine adjacent portions with matching formatting, and reapply the merged runs to the shape. | Explain how to extend the merge logic to include underline, strikeout, superscript, and subscript properties. | Provide a GitHub‑style README snippet describing this example and its prerequisites.

using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsRichTextMergeDemo
{
    // This Aspose.Cells for .NET example creates a rectangle shape, adds rich‑text with alternating bold/normal segments, extracts FontSetting runs via GetRichFormattings, merges consecutive runs that share the same Font and TextOptions, reapplies the consolidated formatting, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape with rich text containing different formatting parts
            // Parameters: upper left row, upper left column, upper left offset, height, width, lower right offset
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
            shape.Text = "Bold text. Normal text. Bold text again. Normal again.";

            // Apply formatting to create separate portions
            // Portion 1: "Bold text." (bold)
            shape.Characters(0, 10).Font.IsBold = true;

            // Portion 2: " Normal text. " (normal)
            shape.Characters(10, 13).Font.IsBold = false;

            // Portion 3: "Bold text again." (bold)
            shape.Characters(23, 16).Font.IsBold = true;

            // Portion 4: " Normal again." (normal)
            shape.Characters(39, 13).Font.IsBold = false;

            // Retrieve the existing rich text portions
            FontSetting[] portions = shape.GetRichFormattings();

            // List to hold merged portions (start index, length, font, text options)
            var mergedPortions = new List<(int Start, int Length, Font Font, TextOptions Options)>();
            var mergedTextBuilder = new StringBuilder();

            foreach (FontSetting setting in portions)
            {
                // Extract the text for the current portion
                string partText = shape.Text.Substring(setting.StartIndex, setting.Length);

                // Capture formatting details
                Font currentFont = setting.Font;
                TextOptions currentOptions = setting.TextOptions;

                // Determine if we can merge with the previous portion
                bool canMerge = false;
                if (mergedPortions.Count > 0)
                {
                    var last = mergedPortions[mergedPortions.Count - 1];

                    // Compare essential formatting properties (you can extend this comparison as needed)
                    bool sameFont = last.Font.IsBold == currentFont.IsBold &&
                                    last.Font.IsItalic == currentFont.IsItalic &&
                                    last.Font.Size == currentFont.Size &&
                                    last.Font.Color.ToArgb() == currentFont.Color.ToArgb();

                    bool sameOptions = last.Options.IsBold == currentOptions.IsBold &&
                                       last.Options.IsItalic == currentOptions.IsItalic;

                    canMerge = sameFont && sameOptions;
                }

                if (canMerge)
                {
                    // Extend the previous merged portion
                    var last = mergedPortions[mergedPortions.Count - 1];
                    mergedPortions[mergedPortions.Count - 1] = (last.Start, last.Length + partText.Length, last.Font, last.Options);
                }
                else
                {
                    // Create a new merged portion entry
                    mergedPortions.Add((mergedTextBuilder.Length, partText.Length, currentFont, currentOptions));
                }

                // Append the text to the combined string
                mergedTextBuilder.Append(partText);
            }

            // Replace the shape's text with the merged text
            shape.Text = mergedTextBuilder.ToString();

            // Reapply formatting based on the merged portions
            foreach (var mp in mergedPortions)
            {
                var chars = shape.Characters(mp.Start, mp.Length);
                // Apply font formatting
                chars.Font.IsBold = mp.Font.IsBold;
                chars.Font.IsItalic = mp.Font.IsItalic;
                chars.Font.Size = mp.Font.Size;
                chars.Font.Color = mp.Font.Color;

                // Apply text options (e.g., bold/italic via TextOptions if needed)
                chars.TextOptions.IsBold = mp.Options.IsBold;
                chars.TextOptions.IsItalic = mp.Options.IsItalic;
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RichTextPortionMergeDemo.xlsx");
        }
    }
}
