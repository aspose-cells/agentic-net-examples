// Title: Group three WordArt shapes, lock the group, and save the worksheet as an XLSX file using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds three WordArt shapes to a worksheet, sets each shape's font size, groups them into a GroupShape, locks the group, and saves the workbook as an XLSX file. | Explain the steps to lock a GroupShape containing WordArt objects before exporting the workbook with Aspose.Cells for .NET. | Show how to adjust WordArt positioning, apply a preset style, and protect the grouped shapes when saving an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# group WordArt shapes and lock them before saving | how to lock a GroupShape with WordArt in an Excel workbook using Aspose.Cells | saving an Excel file with grouped WordArt objects as XLSX in .NET | set font size for multiple WordArt shapes with Aspose.Cells
// Tags: group WordArt shapes Aspose.Cells .NET | lock GroupShape Aspose.Cells | save workbook as XLSX with grouped shapes | set WordArt font size Aspose.Cells | add WordArt to worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts three WordArt shapes with distinct preset styles and positions, sets each shape's font size, groups them into a single GroupShape, locks the group to prevent editing, and saves the workbook as 'GroupedWordArt.xlsx' in XLSX format.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add three WordArt shapes with different positions and texts (using positional arguments)
            Shape wordArt1 = sheet.Shapes.AddWordArt(
                (PresetWordArtStyle)0,               // First preset style
                "First WordArt",
                2,    // upperLeftRow
                2,    // upperLeftColumn
                20,   // top (or row offset depending on API version)
                20,   // left (or column offset depending on API version)
                100,  // height
                200); // width

            Shape wordArt2 = sheet.Shapes.AddWordArt(
                (PresetWordArtStyle)1,               // Second preset style
                "Second WordArt",
                5,    // upperLeftRow
                5,    // upperLeftColumn
                120,  // top / row offset
                20,   // left / column offset
                100,
                200);

            Shape wordArt3 = sheet.Shapes.AddWordArt(
                (PresetWordArtStyle)2,               // Third preset style
                "Third WordArt",
                8,    // upperLeftRow
                8,    // upperLeftColumn
                220,  // top / row offset
                20,   // left / column offset
                100,
                200);

            // Set font size for each WordArt
            wordArt1.Font.Size = 24;
            wordArt2.Font.Size = 24;
            wordArt3.Font.Size = 24;

            // Group the three WordArt shapes
            Shape[] shapesToGroup = { wordArt1, wordArt2, wordArt3 };
            GroupShape group = sheet.Shapes.Group(shapesToGroup);

            // Save the workbook as an XLSX file
            const string outputPath = "GroupedWordArt.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
