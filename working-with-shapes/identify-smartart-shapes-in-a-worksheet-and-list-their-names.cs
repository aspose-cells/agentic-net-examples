using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtIdentifier
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    // Output the name of the SmartArt shape
                    Console.WriteLine($"SmartArt Shape Name: {shape.Name}");
                }
            }

            // Optionally save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}