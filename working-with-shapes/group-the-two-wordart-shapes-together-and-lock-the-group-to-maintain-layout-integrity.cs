// Title: Group two WordArt shapes and lock the group in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds two WordArt objects to a worksheet, merges them into a single group, sets the group to be locked, and saves the workbook with Aspose.Cells. | Show how to create WordArt shapes, combine them into a grouped shape, prevent the group from being moved, and export the file using Aspose.Cells in a .NET application. | Demonstrate locking a collection of WordArt shapes after grouping them so the layout stays fixed when the workbook is opened, using Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# how to group WordArt shapes and lock them | prevent grouped WordArt from moving in Excel with Aspose.Cells .NET | example of locking a shape group in an Excel file using Aspose.Cells | C# code to combine WordArt objects into a locked group with Aspose.Cells
// Tags: Aspose.Cells shape grouping C# | Excel WordArt lock group .NET | prevent shape repositioning Aspose.Cells | grouping WordArt objects Aspose.Cells | lock shape collection Excel C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates two WordArt shapes on the first worksheet, groups them into a single shape, locks the group to keep its layout fixed, and saves the workbook using Aspose.Cells for .NET.
class GroupWordArtExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; otherwise create a new empty workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Add first WordArt shape (style, text, row, column, top, left, height, width)
            Shape wordArt1 = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Hello",
                0, 0,
                50, 50,
                60, 150);
            wordArt1.Font.Color = Color.Blue;

            // Add second WordArt shape
            Shape wordArt2 = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "World",
                0, 0,
                150, 50,
                60, 150);
            wordArt2.Font.Color = Color.Red;

            // Group the two WordArt shapes together
            Shape[] shapesToGroup = new Shape[] { wordArt1, wordArt2 };
            Shape groupShape = sheet.Shapes.Group(shapesToGroup);

            // Lock the group to maintain layout integrity
            groupShape.IsLocked = true;

            // Save the workbook with the grouped and locked WordArt shapes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
