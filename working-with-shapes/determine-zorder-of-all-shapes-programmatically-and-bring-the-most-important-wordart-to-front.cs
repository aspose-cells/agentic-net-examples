// Title: Read Z‑order of every shape in an Excel worksheet and bring the highest‑layer WordArt to the front with Aspose.Cells for .NET
// AI Prompts: List each worksheet shape's stacking index and display its name, type, and order using Aspose.Cells. | Select the WordArt shape with the greatest stacking index and set its index to the topmost value so it appears above all other shapes. | Save the modified workbook to a new file after reordering shapes, handling the case where the source file is missing.
// Common Searches: Aspose.Cells C# get Z‑order of shapes in a worksheet | C# move WordArt to front in Excel file using Aspose.Cells | list all shape properties including ZOrderPosition with Aspose.Cells .NET | how to programmatically bring a specific shape to the top layer in Aspose.Cells | determine which WordArt is on top of other shapes in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition enumeration | Aspose.Cells bring WordArt to front | Aspose.Cells identify WordArt shapes | Aspose.Cells reorder worksheet shapes | Aspose.Cells save workbook after shape manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through all shapes on the first worksheet printing each shape’s name, type, and Z‑order position, finds the WordArt shape with the highest Z‑order, assigns it a ZOrderPosition equal to the total shape count to bring it to the front, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the sheet
            ShapeCollection shapes = sheet.Shapes;

            // Output Z‑order of all shapes
            Console.WriteLine("Z‑order of all shapes:");
            foreach (Shape shape in shapes)
            {
                Console.WriteLine($"Name: {shape.Name}, Type: {shape.Type}, ZOrder: {shape.ZOrderPosition}");
            }

            // Find the most important WordArt (the one with the highest ZOrder)
            Shape mostImportantWordArt = null;
            int maxZOrder = int.MinValue;

            foreach (Shape shape in shapes)
            {
                // Use IsWordArt property to identify WordArt shapes
                if (shape.IsWordArt)
                {
                    if (shape.ZOrderPosition > maxZOrder)
                    {
                        maxZOrder = shape.ZOrderPosition;
                        mostImportantWordArt = shape;
                    }
                }
            }

            // Bring the selected WordArt to the front by setting its Z‑order to the highest value
            if (mostImportantWordArt != null)
            {
                try
                {
                    // Assign a Z‑order greater than any existing shape
                    mostImportantWordArt.ZOrderPosition = shapes.Count;
                    Console.WriteLine($"WordArt \"{mostImportantWordArt.Name}\" brought to front.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to bring WordArt to front: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
