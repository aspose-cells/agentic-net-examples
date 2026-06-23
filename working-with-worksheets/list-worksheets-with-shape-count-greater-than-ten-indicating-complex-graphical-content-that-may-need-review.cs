using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file (adjust the path as needed)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate over all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the number of drawing shapes on the current worksheet
            int shapeCount = sheet.Shapes.Count;

            // If the worksheet contains more than ten shapes, output its name and count
            if (shapeCount > 10)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has {shapeCount} shapes (complex graphical content).");
            }
        }
    }
}