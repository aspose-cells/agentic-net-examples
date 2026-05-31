using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or iterate through all worksheets as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Check if the shape is a SmartArt object
            if (shape.IsSmartArt)
            {
                // Convert the SmartArt to a grouped shape
                GroupShape groupShape = shape.GetResultOfSmartArt();

                if (groupShape != null)
                {
                    // Iterate through each child shape within the grouped SmartArt
                    foreach (Shape childShape in groupShape.GetGroupedShapes())
                    {
                        // Read the Text property of the child shape
                        string childText = childShape.Text;

                        // Output the extracted text (you can process it further as needed)
                        Console.WriteLine($"Child Shape ID: {childShape.Id}, Text: {childText}");
                    }
                }
            }
        }

        // Save the workbook (optional, if you need to persist any changes)
        workbook.Save("output.xlsx");
    }
}