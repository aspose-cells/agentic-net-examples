// Title: Extract font colors of each RichTextPortion from shapes, cells, and comments using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate over FontSetting objects from a shape, a worksheet cell, and a comment, building a Dictionary<int, Color> that maps each RichTextPortion start index to its Font.Color. | Show how to call GetRichFormattings and GetCharacters in Aspose.Cells, collect the font colors of rich‑text segments, store them in a dictionary, and output the mappings to the console.
// Common Searches: Aspose.Cells C# get font color of each rich text portion in a shape | Retrieve colors of rich text segments in an Excel cell using Aspose.Cells .NET | How to read comment rich text font colors with Aspose.Cells and store in a dictionary | Example of extracting RichTextPortion colors from shapes, cells, and comments in Aspose.Cells
// Tags: Aspose.Cells GetRichFormattings font color extraction | C# dictionary rich text portion colors | shape rich text color mapping Aspose.Cells | cell rich text font color Aspose.Cells | comment rich text color extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRichTextColorExtraction
{
    // The example creates a workbook, adds a rectangle shape, a cell, and a comment each containing colored text fragments. It uses GetRichFormattings (for shapes and comments) and GetCharacters (for cells) to obtain FontSetting arrays, then iterates through them to populate dictionaries that map each RichTextPortion's start index to its Font.Color. The mappings are printed to the console and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Example 1: Shape with rich text portions
            // -------------------------------------------------
            // Add a rectangle shape
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
            shape.Text = "RedBlueGreen";

            // Apply different colors to portions of the shape text
            shape.Characters(0, 3).Font.Color = Color.Red;      // "Red"
            shape.Characters(3, 4).Font.Color = Color.Blue;     // "Blue"
            shape.Characters(7, 5).Font.Color = Color.Green;    // "Green"

            // Retrieve rich text formatting (FontSetting objects)
            FontSetting[] shapeFormattings = shape.GetRichFormattings();

            // Store colors in a dictionary: key = start index, value = font color
            Dictionary<int, Color> shapeColors = new Dictionary<int, Color>();
            foreach (FontSetting setting in shapeFormattings)
            {
                shapeColors[setting.StartIndex] = setting.Font.Color;
            }

            // -------------------------------------------------
            // Example 2: Cell with rich text portions
            // -------------------------------------------------
            // Set HTML string with colored spans
            worksheet.Cells["A1"].HtmlString = "<font color='Red'>Red</font><font color='Blue'>Blue</font><font color='Green'>Green</font>";

            // Get character formatting for the cell
            FontSetting[] cellFormattings = worksheet.Cells["A1"].GetCharacters();

            // Store colors in a dictionary: key = start index, value = font color
            Dictionary<int, Color> cellColors = new Dictionary<int, Color>();
            foreach (FontSetting setting in cellFormattings)
            {
                cellColors[setting.StartIndex] = setting.Font.Color;
            }

            // -------------------------------------------------
            // Example 3: Comment with rich text portions
            // -------------------------------------------------
            // Add a comment to cell B2
            int commentIndex = worksheet.Comments.Add("B2");
            Comment comment = worksheet.Comments[commentIndex];
            comment.HtmlNote = "<font color='Red'>Red</font> <font color='Blue'>Blue</font> <font color='Green'>Green</font>";

            // Get rich text formatting for the comment
            FontSetting[] commentFormattings = comment.GetRichFormattings();

            // Store colors in a dictionary: key = start index, value = font color
            Dictionary<int, Color> commentColors = new Dictionary<int, Color>();
            foreach (FontSetting setting in commentFormattings)
            {
                commentColors[setting.StartIndex] = setting.Font.Color;
            }

            // -------------------------------------------------
            // Output the collected colors to console
            // -------------------------------------------------
            Console.WriteLine("Shape Rich Text Portion Colors:");
            foreach (var kvp in shapeColors)
            {
                Console.WriteLine($"StartIndex: {kvp.Key}, Color: {kvp.Value}");
            }

            Console.WriteLine("\nCell Rich Text Portion Colors:");
            foreach (var kvp in cellColors)
            {
                Console.WriteLine($"StartIndex: {kvp.Key}, Color: {kvp.Value}");
            }

            Console.WriteLine("\nComment Rich Text Portion Colors:");
            foreach (var kvp in commentColors)
            {
                Console.WriteLine($"StartIndex: {kvp.Key}, Color: {kvp.Value}");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("RichTextColorsDemo.xlsx");
        }
    }
}
