// Title: Send the last shape in an Excel worksheet to the back of the Z‑order and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that loads an Excel file, moves the last shape in the first worksheet to the back of the Z‑order, and saves the result to a new file. | Provide a C# example that checks for shapes on a worksheet, sets the ZOrderPosition of the final shape to 0, and writes the workbook to disk with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to change ZOrderPosition of a shape in Excel | move last shape to back of drawing order in Aspose.Cells worksheet | C# code to send a shape behind all others using Aspose.Cells | save workbook after reordering shapes with Aspose.Cells .NET
// Tags: Aspose.Cells set shape ZOrderPosition | C# move Excel shape to back | Aspose.Cells reorder worksheet shapes | save workbook after shape manipulation Aspose.Cells | Excel shape drawing order .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// // Loads an existing Excel file (or creates a new workbook), accesses the first worksheet's ShapeCollection, sets the ZOrderPosition of the last shape to 0 to move it behind all other shapes, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet.
            ShapeCollection shapes = sheet.Shapes;

            // If there is at least one shape, send the last one to the back by setting its Z-order position.
            if (shapes.Count > 0)
            {
                Shape lastShape = shapes[shapes.Count - 1];
                // Setting ZOrderPosition to 0 moves the shape behind all others.
                lastShape.ZOrderPosition = 0;
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
