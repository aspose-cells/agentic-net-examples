using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class RichTextPortionSummary
    {
        public static void Main()
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
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Sample data: add rich text to a cell, a shape and a comment
            // -------------------------------------------------
            // Cell with rich text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Hello World");
            // Apply formatting to "Hello"
            FontSetting cellPart1 = cell.Characters(0, 5);
            cellPart1.Font.IsBold = true;
            cellPart1.Font.Name = "Arial";
            // Apply formatting to "World"
            FontSetting cellPart2 = cell.Characters(6, 5);
            cellPart2.Font.IsItalic = true;
            cellPart2.Font.Name = "Calibri";

            // Shape with rich text
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 100, 200, 0);
            shape.Text = "Shape Text Example";
            shape.Characters(0, 5).Font.Name = "Times New Roman";
            shape.Characters(6, 4).Font.Name = "Verdana";

            // Comment with rich text
            int commentIndex = worksheet.Comments.Add("B2");
            Comment comment = worksheet.Comments[commentIndex];
            comment.HtmlNote = "<b>Bold</b> and <i>Italic</i>";
            // Note: formatting is already embedded in the HTML

            // -------------------------------------------------
            // Summarize RichTextPortion (FontSetting) information
            // -------------------------------------------------
            Console.WriteLine("=== Rich Text Portions in Worksheet ===");

            // 1. Cells
            Console.WriteLine("\n--- Cells ---");
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell curCell = worksheet.Cells[row, col];
                    if (curCell == null) continue;

                    // Check if the cell contains rich text
                    if (curCell.IsRichText())
                    {
                        FontSetting[] settings = curCell.GetCharacters();
                        foreach (FontSetting fs in settings)
                        {
                            Console.WriteLine($"Cell {curCell.Name}: StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                        }
                    }
                }
            }

            // 2. Shapes
            Console.WriteLine("\n--- Shapes ---");
            foreach (Shape shp in worksheet.Shapes)
            {
                if (shp.IsRichText)
                {
                    FontSetting[] settings = shp.GetRichFormattings();
                    foreach (FontSetting fs in settings)
                    {
                        // Shape does not expose an Index property; use Name instead
                        Console.WriteLine($"Shape (Name \"{shp.Name}\"): StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                    }
                }
            }

            // 3. Comments
            Console.WriteLine("\n--- Comments ---");
            foreach (Comment cmt in worksheet.Comments)
            {
                // Comments always support rich text formatting
                FontSetting[] settings = cmt.GetRichFormattings();
                foreach (FontSetting fs in settings)
                {
                    // Use the cell reference stored in the comment's Note property if needed
                    Console.WriteLine($"Comment on {cmt.Note}: StartIndex={fs.StartIndex}, Length={fs.Length}, FontName={fs.Font.Name}");
                }
            }

            // Save the workbook (lifecycle: save)
            string outputPath = "RichTextPortionSummary.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to \"{Path.GetFullPath(outputPath)}\"");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}