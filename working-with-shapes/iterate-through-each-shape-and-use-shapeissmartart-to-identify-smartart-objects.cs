using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class IdentifySmartArt
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through each shape in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Use the IsSmartArt property to check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    Console.WriteLine($"SmartArt found in worksheet '{worksheet.Name}' with shape name '{shape.Name}'.");
                }
            }
        }

        // Save the workbook (no modifications made, just demonstrating lifecycle)
        workbook.Save("output.xlsx");
    }
}