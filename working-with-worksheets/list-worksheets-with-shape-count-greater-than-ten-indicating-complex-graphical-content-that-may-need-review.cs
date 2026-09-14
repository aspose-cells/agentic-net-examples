// Title: Identify Excel worksheets containing more than ten shapes with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that returns a list of worksheet names where the number of shapes exceeds ten. | Create a method that scans an Excel workbook with Aspose.Cells and logs each sheet's name and shape count when the count is greater than ten. | Generate a console application in C# that loads a workbook, iterates worksheets, and prints sheets that contain over ten graphical objects.
// Common Searches: aspocells c# find worksheets with many shapes | how to list Excel sheets that have over 10 drawings using Aspose.Cells | C# count shapes in each worksheet Aspose.Cells example | filter workbook worksheets by number of shapes .NET | identify Excel worksheets with complex graphics Aspose.Cells
// Tags: Aspose.Cells worksheet shape enumeration | filter Excel sheets by shape count .NET | detect high shape density in workbook | list worksheets with extensive graphics Aspose.Cells | C# shape count threshold for Excel worksheets

using System;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates through every worksheet, counts the shapes on each sheet, and outputs the names of worksheets that have more than ten shapes, highlighting sheets with complex graphical content.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Count the number of shapes on the current worksheet
            int shapeCount = sheet.Shapes.Count;

            // If the worksheet has more than ten shapes, output its name
            if (shapeCount > 10)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has {shapeCount} shapes (complex graphical content).");
            }
        }
    }
}
