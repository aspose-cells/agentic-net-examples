// Title: Align all worksheet shapes to the printable area's left margin using Aspose.Cells for .NET
// AI Prompts: Set each shape's Left property to the worksheet's left margin value with Aspose.Cells in C#. | Iterate over Worksheet.Shapes and align them to the printable left edge based on PageSetup. | Modify an existing workbook so that all drawing objects start at the page's left margin.
// Common Searches: C# Aspose.Cells move all shapes to the printable left margin | How to align Excel shapes with page margins using Aspose.Cells | Set shape left position from worksheet PageSetup margins in .NET | Batch reposition worksheet drawings to the left printable area with Aspose.Cells
// Tags: shape left alignment with PageSetup margin | Aspose.Cells batch shape repositioning | worksheet drawing objects left margin adjustment | C# align Excel shapes to printable area | Aspose.Cells set shape.Left from worksheet margins

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads a workbook, reads the first worksheet's left margin from PageSetup, loops through all shapes on that sheet, sets each shape's Left property to the rounded left‑margin value, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the left margin of the printable area (in points)
            double printableLeft = sheet.PageSetup.LeftMargin;

            // Align each shape in the worksheet to the left edge of the printable area
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Shape.Left expects an integer value; convert from double safely
                    shape.Left = (int)Math.Round(printableLeft);
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Failed to align shape '{shape.Name}': {exShape.Message}");
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
