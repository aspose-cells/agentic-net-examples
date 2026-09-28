// Title: Insert WordArt into a worksheet range and anchor it to move and resize with cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a WordArt graphic to cells B2:D5, sets UpperLeftRow/Column and LowerRightRow/Column, and configures Placement = MoveAndSize with Aspose.Cells. | Create a C# example that builds a workbook, inserts a WordArt element with custom text, anchors it to a defined cell range, and saves the workbook as an .xlsx file.
// Common Searches: asp.net add WordArt graphic to specific cell range using Aspose.Cells | how to anchor a WordArt object to cells so it moves with rows and columns in Aspose.Cells C# | set placement to MoveAndSize for WordArt in Aspose.Cells workbook | Aspose.Cells C# example for positioning WordArt between B2 and D5
// Tags: WordArt insertion Aspose.Cells | WordArt placement MoveAndSize | anchor graphic to cell range C# | C# Aspose.Cells shape positioning | save workbook with WordArt Xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a WordArt element with the text "Aspose" into the B2:D5 range, anchors it by setting UpperLeftRow/Column and LowerRightRow/Column, sets the placement to MoveAndSize so the graphic moves and resizes with the cells, and saves the file as WordArtExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range to anchor the WordArt (e.g., B2:D5)
            int startRow = 1;      // Row index for B2 (zero‑based)
            int startColumn = 1;   // Column index for B2
            int endRow = 4;        // Row index for D5
            int endColumn = 3;     // Column index for D5

            // Add a WordArt shape with a preset style
            // Parameters: preset style, text, upper‑left row, upper‑left column,
            // row offset, column offset, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Aspose",
                startRow,
                startColumn,
                0,
                0,
                200,
                50);

            // Anchor the shape to the defined cell range
            wordArt.UpperLeftRow = startRow;
            wordArt.UpperLeftColumn = startColumn;
            wordArt.LowerRightRow = endRow;
            wordArt.LowerRightColumn = endColumn;

            // Ensure the shape moves and resizes with the cells
            wordArt.Placement = PlacementType.MoveAndSize;

            // Define output file path
            string outputPath = "WordArtExample.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
