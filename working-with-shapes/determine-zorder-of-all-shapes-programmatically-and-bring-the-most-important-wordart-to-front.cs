using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes to display their Z‑order positions
            Console.WriteLine("Current Z‑order of shapes:");
            foreach (Shape shape in sheet.Shapes)
            {
                Console.WriteLine($"Shape Name: {shape.Name}, ZOrderPosition: {shape.ZOrderPosition}");
            }

            // Identify the most important WordArt.
            // For demonstration, we assume the WordArt shape's name contains "WordArt".
            Shape mostImportantWordArt = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name != null && shape.Name.IndexOf("WordArt", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Choose the WordArt with the highest current Z‑order (largest ZOrderPosition value)
                    if (mostImportantWordArt == null || shape.ZOrderPosition > mostImportantWordArt.ZOrderPosition)
                    {
                        mostImportantWordArt = shape;
                    }
                }
            }

            if (mostImportantWordArt != null)
            {
                // Bring the selected WordArt to the front (ZOrderPosition = 0)
                mostImportantWordArt.ZOrderPosition = 0;

                // Alternatively, you could use ToFrontOrBack with a positive value:
                // mostImportantWordArt.ToFrontOrBack(1);
                Console.WriteLine($"Moved WordArt '{mostImportantWordArt.Name}' to the front.");
            }
            else
            {
                Console.WriteLine("No WordArt shape found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}