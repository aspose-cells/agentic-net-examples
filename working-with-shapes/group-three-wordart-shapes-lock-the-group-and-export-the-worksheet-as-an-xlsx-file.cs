// Title: Group and Lock WordArt Shapes, Then Save as XLSX with Aspose.Cells for .NET (C#)
// Description: A concise C# example that creates a workbook, adds three WordArt shapes with different preset styles, groups them into a single GroupShape, locks the group (effective after worksheet protection), and saves the result as an XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells C# | WordArt grouping | GroupShape lock | IsLocked property | save workbook as XLSX | shape collection Aspose | preset WordArt style | protect worksheet Aspose.Cells | example code GitHub | Aspose.Cells shape manipulation
// Common Searches: how to group WordArt in Aspose.Cells .NET | lock grouped shapes before protecting worksheet Aspose.Cells | save grouped WordArt to XLSX with C# | Aspose.Cells GroupShape example | C# code to add WordArt and lock group
// Developer Intent: Create three WordArt objects, combine them into a locked group, and export the worksheet as an XLSX file.
// Use Cases: Design a composite title banner that can be moved or resized as a single unit. | Prevent end‑users from editing or repositioning critical WordArt headings by locking the group before worksheet protection. | Generate a styled report template with grouped WordArt elements and distribute it as an XLSX file.
// AI Prompts: Generate C# code with Aspose.Cells that adds multiple WordArt shapes, groups them, sets IsLocked = true, and saves the workbook as XLSX. | Explain the interaction between GroupShape.IsLocked and worksheet protection in Aspose.Cells. | Provide a variation that groups WordArt, locks the group, and writes the workbook to a MemoryStream instead of a file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // A concise C# example that creates a workbook, adds three WordArt shapes with different preset styles, groups them into a single GroupShape, locks the group (effective after worksheet protection), and saves the result as an XLSX file using Aspose.Cells.
    public class GroupWordArtDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                ShapeCollection shapes = sheet.Shapes;

                // Add three WordArt shapes with different preset styles
                Shape wordArt1 = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "First",
                    topRow: 2, top: 10,
                    leftColumn: 2, left: 10,
                    height: 50, width: 200);

                Shape wordArt2 = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2,
                    "Second",
                    topRow: 5, top: 10,
                    leftColumn: 5, left: 10,
                    height: 50, width: 200);

                Shape wordArt3 = shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle3,
                    "Third",
                    topRow: 8, top: 10,
                    leftColumn: 8, left: 10,
                    height: 50, width: 200);

                // Group the three WordArt shapes
                Shape[] wordArts = new Shape[] { wordArt1, wordArt2, wordArt3 };
                GroupShape group = shapes.Group(wordArts);

                // Lock the group (effective when the worksheet is protected)
                group.IsLocked = true;

                // Save the workbook as an XLSX file
                workbook.Save("GroupedWordArt.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GroupWordArtDemo.Run();
        }
    }
}
