// Title: How to read a shape’s absolute position, apply a random offset, and save the updated location in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# get worksheet shape top and left values with Aspose.Cells | apply random movement to an Excel drawing object using Aspose.Cells .NET | programmatically shift shape position in an .xlsx workbook Aspose.Cells | retrieve and modify shape coordinates in Aspose.Cells C# example | move first shape on a worksheet by a random offset with Aspose.Cells
// Tags: Aspose.Cells shape coordinate manipulation | C# random offset for worksheet shape | update shape Top Left properties Aspose.Cells | Excel drawing object repositioning .NET | absolute shape position retrieval Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, checks for at least one shape on the first worksheet, reads the shape’s absolute Top and Left coordinates (in points), generates a random offset between –10 and +10 points for both axes, applies the offset to the shape’s position, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Retrieve the first shape on the worksheet
            Shape shape = worksheet.Shapes[0];

            // Get the shape's current absolute position in points (int values)
            int currentTop = shape.Top;   // Distance from the top edge of the worksheet
            int currentLeft = shape.Left; // Distance from the left edge of the worksheet

            // Create a random offset between -10 and +10 points for both axes
            Random random = new Random();
            int offsetTop = (int)(random.NextDouble() * 20 - 10);   // -10 .. +10
            int offsetLeft = (int)(random.NextDouble() * 20 - 10);  // -10 .. +10

            // Apply the random offset to the shape's position
            shape.Top = currentTop + offsetTop;
            shape.Left = currentLeft + offsetLeft;

            // Save the workbook with the updated shape location
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
