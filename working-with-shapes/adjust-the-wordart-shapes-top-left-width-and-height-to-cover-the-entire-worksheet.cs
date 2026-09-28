// Title: How to resize an Aspose.Cells WordArt shape to fill the entire worksheet in C#
// AI Prompts: Create a WordArt shape on a worksheet, compute the total pixel width of all used columns and the total pixel height of all used rows, then set the shape's Top, Left, Width, and Height properties so it spans the whole sheet using Aspose.Cells for .NET. | Programmatically adjust a WordArt object's dimensions to match the used range of an Excel worksheet by iterating through columns and rows to obtain pixel sizes with GetColumnWidthPixel and GetRowHeightPixel in C#.
// Common Searches: Aspose.Cells C# set WordArt size to match used range | Resize WordArt to cover full worksheet with Aspose.Cells | Calculate pixel dimensions of Excel sheet for shape sizing in .NET | How to make WordArt fill entire Excel sheet using Aspose.Cells API | C# code to expand WordArt across all rows and columns in a workbook
// Tags: Aspose.Cells WordArt full sheet sizing | C# compute worksheet pixel dimensions Aspose.Cells | adjust WordArt size to used range Aspose.Cells | shape spanning entire worksheet Aspose.Cells | Excel WordArt dimension control C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt shape, calculates the combined pixel width of all used columns and the combined pixel height of all used rows, positions the shape at the top‑left corner, assigns the calculated Width and Height so the WordArt covers the entire worksheet area, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape to the worksheet
            // Parameters: style, text, font size, top, left, width, height, rotationAngle
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Full Sheet WordArt",
                24,
                0, 0, 0, 0, 0);

            // Calculate total width of all used columns (in pixels)
            int totalWidth = 0;
            int maxColumn = worksheet.Cells.MaxColumn; // zero‑based index of the last used column
            for (int col = 0; col <= maxColumn; col++)
            {
                totalWidth += (int)worksheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate total height of all used rows (in pixels)
            int totalHeight = 0;
            int maxRow = worksheet.Cells.MaxRow; // zero‑based index of the last used row
            for (int row = 0; row <= maxRow; row++)
            {
                totalHeight += (int)worksheet.Cells.GetRowHeightPixel(row);
            }

            // Position the WordArt at the top‑left corner (cell A1)
            wordArt.Top = 0;
            wordArt.Left = 0;

            // Resize the WordArt to cover the entire worksheet area
            wordArt.Width = totalWidth;
            wordArt.Height = totalHeight;

            // Save the workbook
            string outputPath = "WordArtFullSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
