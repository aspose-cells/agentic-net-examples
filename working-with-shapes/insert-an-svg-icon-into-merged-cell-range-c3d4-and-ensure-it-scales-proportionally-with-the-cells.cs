// Title: How to add an SVG icon to a merged cell range (C3:D4) and keep it proportional with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert SVG into merged cells and maintain aspect ratio | scale picture with MoveAndSize placement in a merged range using Aspose.Cells | add SVG file to cell range C3:D4 in Excel with Aspose.Cells .NET | merge cells and embed scalable SVG icon using Aspose.Cells API | C# Aspose.Cells picture placement type MoveAndSize for merged cells
// Tags: svg picture insertion merged cells Aspose.Cells | picture placement scaling merged cells | c# aspnet aspose.cells add picture to range | excel workbook embed scalable svg | merged cell range picture placement .net

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells C3:D4, checks for an SVG file, inserts the SVG into the merged range with Pictures.Add, sets the picture's Placement to MoveAndSize so it scales proportionally with the cells, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Merge the cell range C3:D4 (zero‑based indices: row 2, column 2)
            sheet.Cells.Merge(2, 2, 2, 2);

            // Path to the SVG icon file
            string svgPath = "icon.svg";

            // Verify that the SVG file exists before attempting to load it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Insert the SVG picture into the merged range.
            // Add returns the index of the picture in the collection.
            int pictureIndex = sheet.Pictures.Add(2, 2, 3, 3, svgPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Resize the picture together with the cells
            picture.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
