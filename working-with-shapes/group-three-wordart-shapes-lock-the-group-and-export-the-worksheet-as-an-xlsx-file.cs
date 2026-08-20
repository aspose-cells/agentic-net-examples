// Title: Group and Lock WordArt Shapes, then Export as XLSX with Aspose.Cells for .NET
// Description: Demonstrates how to create three WordArt shapes, combine them into a GroupShape, set the group as locked, and save the workbook as an XLSX file using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells | C# | .NET | WordArt | AddWordArt | GroupShape | lock shapes | IsLocked property | protect worksheet | export XLSX | worksheet shapes
// Common Searches: Aspose.Cells group WordArt C# | lock grouped shapes Aspose.Cells | save workbook as XLSX after grouping shapes | how to protect shape groups in Aspose.Cells | AddWordArt and GroupShape example
// Developer Intent: Create three WordArt objects, group them, lock the group, and save the worksheet as an XLSX file.
// Use Cases: Design a fixed header with styled WordArt that stays in place when the sheet is protected. | Build a reusable template where grouped WordArt titles cannot be moved or edited by end users. | Generate a report footer with locked WordArt to ensure consistent branding across shared workbooks.
// AI Prompts: Generate C# code that adds four WordArt shapes, groups them, locks the group, and saves the workbook as XLSX using Aspose.Cells. | Explain how to unlock and ungroup a locked WordArt group in an existing Aspose.Cells workbook. | Provide robust error handling for grouping WordArt when the worksheet already contains many shapes or unsupported shape types.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupWordArtDemo
{
    // Demonstrates how to create three WordArt shapes, combine them into a GroupShape, set the group as locked, and save the workbook as an XLSX file using Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add three WordArt shapes with different texts
            Shape wordArt1 = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "First",
                2,   // topRow
                10,  // top (pixels)
                2,   // leftColumn
                10,  // left (pixels)
                50,  // height (pixels)
                200  // width (pixels)
            );

            Shape wordArt2 = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle2,
                "Second",
                5,
                10,
                5,
                10,
                50,
                200
            );

            Shape wordArt3 = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle3,
                "Third",
                8,
                10,
                8,
                10,
                50,
                200
            );

            // Group the three WordArt shapes
            Shape[] wordArtArray = new Shape[] { wordArt1, wordArt2, wordArt3 };
            GroupShape group = shapes.Group(wordArtArray);

            // Lock the group so it cannot be modified when the sheet is protected
            group.IsLocked = true;

            // Save the workbook as an XLSX file
            workbook.Save("GroupedWordArt.xlsx");
        }
    }
}
