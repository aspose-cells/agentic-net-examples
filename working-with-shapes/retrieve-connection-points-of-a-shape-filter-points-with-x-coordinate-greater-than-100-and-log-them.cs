// Title: Retrieve and filter shape connection points in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, obtains the connection points of the first shape on the first worksheet, and prints only the points whose X coordinate exceeds 100. | Extend the shape‑connection‑point extraction to also apply a Y‑coordinate filter (e.g., Y < 50) and output the matching coordinates. | Add comprehensive error handling to the Aspose.Cells shape processing routine to gracefully manage missing files, worksheets without shapes, and null or empty connection‑point arrays.
// Common Searches: Aspose.Cells C# get shape connection points from worksheet | filter shape connection points where X > 100 using Aspose.Cells | GetConnectionPoints example code Aspose.Cells for .NET | handle missing shapes when reading Excel with Aspose.Cells | C# retrieve first shape connection points in an Excel file
// Tags: Aspose.Cells GetConnectionPoints C# | filter shape connection points by X coordinate | error handling missing shape Aspose.Cells | load workbook access first shape Aspose.Cells | log connection point coordinates C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads input.xlsx, accesses the first worksheet and its first shape, calls GetConnectionPoints() to obtain X/Y pairs, filters out points with an X value greater than 100, and writes those coordinates to the console while handling missing files, absent shapes, and null data gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Get the first shape (adjust index as needed)
            Shape shape = worksheet.Shapes[0];

            // Retrieve connection points (returns float[][] where each inner array is X,Y)
            float[][] connectionPoints = shape.GetConnectionPoints();

            // If the returned array is null or empty, exit gracefully
            if (connectionPoints == null || connectionPoints.Length == 0)
            {
                Console.WriteLine("The shape does not contain any connection points.");
                return;
            }

            // Iterate through each point (X,Y pair)
            foreach (float[] point in connectionPoints)
            {
                if (point == null || point.Length < 2)
                    continue; // skip malformed entries

                float x = point[0];
                float y = point[1];

                // Filter points with X coordinate greater than 100 and log them
                if (x > 100)
                {
                    Console.WriteLine($"Connection Point - X: {x}, Y: {y}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
