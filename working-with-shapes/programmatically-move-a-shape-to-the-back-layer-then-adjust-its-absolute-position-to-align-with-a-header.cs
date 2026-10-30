// Title: How to send a shape to the back layer and position it at cell A1 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move shape to back layer | Set shape ZOrderPosition to 0 with Aspose.Cells .NET | Align a shape to the top‑left corner of cell A1 using Aspose.Cells | Place an existing shape behind other objects in an Excel worksheet C# | Adjust shape Left and Top coordinates to cell A1 in Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition back layer | Aspose.Cells set shape absolute position | Aspose.Cells align shape to cell A1 | Aspose.Cells move shape behind other objects | Aspose.Cells C# shape layering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook (or creates one), accesses the first worksheet, checks for a shape, sets its ZOrderPosition to 0 to move it to the back layer, sets Left and Top to 0 to align the shape with the top‑left corner of cell A1, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Manipulate the first shape if any exist.
            if (sheet.Shapes.Count > 0)
            {
                try
                {
                    Shape shape = sheet.Shapes[0];

                    // Move the shape to the back layer.
                    shape.ZOrderPosition = 0; // 0 is the backmost position.

                    // Align the shape with the top‑left corner of cell A1.
                    // Since GetPixelLeft/Top are unavailable, use zero offsets.
                    shape.Left = 0;
                    shape.Top = 0;
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Shape processing error: {exShape.Message}");
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
