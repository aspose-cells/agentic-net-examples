// Title: How to list each RichTextPortion’s start index, length, and font name in cells, shapes, and comments using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over all cells, shapes, and comments in a worksheet and prints the start index, length, font name, and text of every RichTextPortion. | Create a method that returns a collection of objects containing the address (or shape name), start index, length, font name, and text for each rich‑text segment in a workbook.
// Common Searches: Aspose.Cells C# get rich text formatting details from a worksheet cell | How to retrieve font names of rich text portions in Excel shapes using Aspose.Cells .NET | Extract start index and length of comment rich text with Aspose.Cells | Iterate through RichTextPortion objects in an Aspose.Cells workbook | List rich text segments and their fonts in an Excel file using Aspose.Cells for .NET
// Tags: Aspose.Cells enumerate RichTextPortion | C# extract rich text font from cells | Aspose.Cells shape rich text formatting | Aspose.Cells comment rich text extraction | RichTextPortion start index length .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsRichTextSummary
{
    // Creates a workbook, adds HTML‑styled rich text to a cell, a rectangle shape, and a comment, then walks through the worksheet to output each RichTextPortion’s start index, length, font name, and the corresponding text for cells, shapes, and comments, finally saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Sample rich text in a cell ----------
                // Set HTML formatted text with different formatting parts
                worksheet.Cells["A1"].HtmlString = "Normal <b>Bold</b> <i>Italic</i> <u>Underline</u>";

                // ---------- Sample rich text in a shape ----------
                // Add a rectangle shape with rich text
                Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 100, 200, 0);
                shape.Text = "Shape <b>Bold</b> Text";

                // ---------- Sample rich text in a comment ----------
                // Add a comment to cell B2 with rich text
                int commentIndex = worksheet.Comments.Add("B2");
                Comment comment = worksheet.Comments[commentIndex];
                comment.HtmlNote = "<b>Bold</b>, <i>Italic</i>, <u>Underline</u>";

                // ----------- Summarize Rich Text Portions in Cells -----------
                Console.WriteLine("Cell Rich Text Portions:");
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        if (cell != null && cell.IsRichText())
                        {
                            FontSetting[] settings = cell.GetCharacters();
                            foreach (FontSetting setting in settings)
                            {
                                string textSegment = cell.StringValue.Substring(setting.StartIndex, setting.Length);
                                string fontName = setting.Font.Name;
                                Console.WriteLine($"Cell {cell.Name}: Start={setting.StartIndex}, Length={setting.Length}, Font=\"{fontName}\", Text=\"{textSegment}\"");
                            }
                        }
                    }
                }

                // ----------- Summarize Rich Text Portions in Shapes -----------
                Console.WriteLine("\nShape Rich Text Portions:");
                foreach (Shape shp in worksheet.Shapes)
                {
                    if (shp.IsRichText)
                    {
                        FontSetting[] settings = shp.GetRichFormattings();
                        foreach (FontSetting setting in settings)
                        {
                            // Shape.Text may contain the full text; extract the segment using start index and length
                            string textSegment = shp.Text.Substring(setting.StartIndex, setting.Length);
                            string fontName = setting.Font.Name;
                            string shapeId = string.IsNullOrEmpty(shp.Name) ? $"Index {shp.Id}" : shp.Name;
                            Console.WriteLine($"Shape ({shapeId}): Start={setting.StartIndex}, Length={setting.Length}, Font=\"{fontName}\", Text=\"{textSegment}\"");
                        }
                    }
                }

                // ----------- Summarize Rich Text Portions in Comments -----------
                Console.WriteLine("\nComment Rich Text Portions:");
                foreach (Comment cmt in worksheet.Comments)
                {
                    FontSetting[] settings = cmt.GetRichFormattings();
                    foreach (FontSetting setting in settings)
                    {
                        // Use plain text representation via Note property
                        string plainText = cmt.Note;
                        // Ensure indices are within bounds
                        if (setting.StartIndex + setting.Length <= plainText.Length)
                        {
                            string textSegment = plainText.Substring(setting.StartIndex, setting.Length);
                            string fontName = setting.Font.Name;
                            // Retrieve cell address of the comment
                            string cellAddress = worksheet.Cells[cmt.Row, cmt.Column].Name;
                            Console.WriteLine($"Comment on {cellAddress}: Start={setting.StartIndex}, Length={setting.Length}, Font=\"{fontName}\", Text=\"{textSegment}\"");
                        }
                    }
                }

                // Save the workbook (optional, just to demonstrate lifecycle)
                workbook.Save("RichTextSummary.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
