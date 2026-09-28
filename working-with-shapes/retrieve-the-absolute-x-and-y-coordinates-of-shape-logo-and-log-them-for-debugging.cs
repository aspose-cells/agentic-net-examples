// Title: Get absolute X and Y coordinates of a named shape (Logo) in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, locates the shape named "Logo" on the first worksheet, and prints its Left and Top values to the console. | Show how to read the absolute X (Left) and Y (Top) coordinates of a specific shape in a workbook for debugging purposes using the Aspose.Cells .NET API.
// Common Searches: how to read the horizontal and vertical coordinates of a shape called Logo in an Excel workbook using Aspose.Cells C# | Aspose.Cells example to output absolute position of a worksheet shape for debugging | C# code to get shape location in points from an .xlsx file with Aspose.Cells | retrieve worksheet shape placement coordinates using Aspose.Cells .NET
// Tags: Aspose.Cells get shape absolute position | C# extract shape left top coordinates | Excel worksheet shape location debugging | named shape coordinate retrieval .NET | Aspose.Cells shape placement logging

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads input.xlsx, accesses the first worksheet, finds the shape named "Logo", reads its absolute X (Left) and Y (Top) coordinates in points, and writes the values to the console for debugging.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape named "Logo"
            Shape logoShape = worksheet.Shapes["Logo"];
            if (logoShape != null)
            {
                // Get absolute X (Left) and Y (Top) coordinates in points
                float x = logoShape.Left;
                float y = logoShape.Top;

                // Output the coordinates
                Console.WriteLine($"Logo shape coordinates - X: {x}, Y: {y}");
            }
            else
            {
                Console.WriteLine("Shape 'Logo' not found.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
