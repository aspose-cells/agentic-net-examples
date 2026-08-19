// Title: C# Example: List RichTextPortion Start Index, Length and Font in Cells, Shapes & Comments with Aspose.Cells
// Description: Demonstrates how to create a workbook, add rich‑text to a cell, a shape and a comment, then iterate through all cells, shapes and comments, detect rich‑text portions, and output each portion's StartIndex, Length and Font.Name. The sample also shows saving the workbook.
// Keywords: Aspose.Cells | C# | RichTextPortion | GetCharacters | GetRichFormattings | font name extraction | start index | text length | cell rich text | shape text formatting | comment rich text | Excel automation example | API usage
// Common Searches: Aspose.Cells list rich text portions C# | How to get start index and font of rich text in Excel using Aspose | Extract rich‑text formatting from cells, shapes and comments | GetCharacters GetRichFormattings Aspose.Cells example | C# code to enumerate rich‑text segments in a workbook
// Developer Intent: Generate a detailed report of every rich‑text segment’s start position, length and font across cells, shapes and comments in an Aspose.Cells workbook.
// Use Cases: Audit formatting consistency in automatically generated Excel reports. | Export a formatting map to CSV or JSON for downstream validation. | Create unit tests that verify specific font styles are applied to designated text ranges.
// AI Prompts: Write a reusable method that returns a list of objects with ElementType, Identifier, StartIndex, Length and FontName for all rich‑text portions in an Aspose.Cells workbook. | Show how to write the rich‑text summary to a CSV file with headers Element, Identifier, StartIndex, Length, FontName. | Provide code to filter the summary to only include portions where the font is bold or the text color is red.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RichTextPortionSummary
{
    // Demonstrates how to create a workbook, add rich‑text to a cell, a shape and a comment, then iterate through all cells, shapes and comments, detect rich‑text portions, and output each portion's StartIndex, Length and Font.Name. The sample also shows saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Sample rich text in a cell ----------
                Cell cell = sheet.Cells["A1"];
                cell.Value = "Hello Aspose.Cells World";

                // Make "Hello" bold
                FontSetting fsHello = cell.Characters(0, 5);
                fsHello.Font.IsBold = true;

                // Make "Aspose.Cells" italic
                FontSetting fsAspose = cell.Characters(6, 13);
                fsAspose.Font.IsItalic = true;

                // ---------- Sample rich text in a shape ----------
                Shape shape = sheet.Shapes.AddRectangle(2, 0, 2, 100, 200, 0);
                shape.Text = "Shape Rich Text Example";

                // Make "Shape" underlined
                shape.Characters(0, 5).Font.Underline = FontUnderlineType.Single;

                // Make "Rich Text" red
                shape.Characters(6, 9).Font.Color = Color.Red;

                // ---------- Sample rich text in a comment ----------
                int commentIdx = sheet.Comments.Add("B2");
                Comment comment = sheet.Comments[commentIdx];
                comment.HtmlNote = "<b>Bold Comment</b> and <i>Italic Comment</i>";

                // Output header
                Console.WriteLine("Rich Text Portion Summary:");
                Console.WriteLine("-----------------------------------");

                // ----- Process cells -----
                foreach (Cell c in sheet.Cells)
                {
                    if (c.IsRichText())
                    {
                        FontSetting[] settings = c.GetCharacters();
                        foreach (FontSetting fs in settings)
                        {
                            Console.WriteLine($"Cell {c.Name}: StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                        }
                    }
                }

                // ----- Process shapes -----
                foreach (Shape s in sheet.Shapes)
                {
                    if (s.IsRichText)
                    {
                        FontSetting[] settings = s.GetRichFormattings();
                        foreach (FontSetting fs in settings)
                        {
                            // Use Shape.Name (or Id) for identification
                            Console.WriteLine($"Shape \"{s.Name}\": StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                        }
                    }
                }

                // ----- Process comments -----
                foreach (Comment cm in sheet.Comments)
                {
                    FontSetting[] settings = cm.GetRichFormattings();
                    foreach (FontSetting fs in settings)
                    {
                        // Use the comment's note text for identification
                        Console.WriteLine($"Comment \"{cm.Note}\": StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                    }
                }

                // Save the workbook (optional)
                workbook.Save("RichTextPortionSummary.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
