// Title: Copy all shapes from one worksheet to another while preserving positions, size, rotation, and text using Aspose.Cells for .NET (C#)
// AI Prompts: Copy every shape from Sheet1 to Sheet2, keeping its cell range, dimensions, rotation angle, and alternative text with Aspose.Cells in C#. | Duplicate Excel text boxes, images, and other drawing objects to a target worksheet while preserving alignment and formatting using the Shapes.AddCopy method. | Programmatically transfer all worksheet shapes to another sheet, retaining height, width, and text properties intact with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy shapes between worksheets preserving coordinates | how to duplicate Excel shapes with same size and rotation using Aspose.Cells | copy text boxes and images from one sheet to another programmatically in .NET
// Tags: Shapes.AddCopy method Aspose.Cells C# | copy worksheet shapes preserving coordinates | duplicate Excel drawing objects Aspose.Cells | preserve shape dimensions and rotation .NET | transfer shape alternative text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook (creating a placeholder if missing), accesses Sheet1 and Sheet2, iterates through each Shape in Sheet1, adds a copy to Sheet2 at the same cell range using Shapes.AddCopy, and then copies height, width, rotation angle, alternative text, and text alignment properties before saving the workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a placeholder if missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                int idx = workbook.Worksheets.Add();
                workbook.Worksheets[idx].Name = "Sheet2";
                workbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created a placeholder workbook at '{inputPath}'.");
            }

            // Source worksheet (contains shapes)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Sheet1' not found.");
                return;
            }

            // Destination worksheet (where shapes will be copied)
            Worksheet destSheet = workbook.Worksheets["Sheet2"];
            if (destSheet == null)
            {
                int index = workbook.Worksheets.Add();
                destSheet = workbook.Worksheets[index];
                destSheet.Name = "Sheet2";
            }

            // Copy each shape from source to destination
            foreach (Shape srcShape in sourceSheet.Shapes)
            {
                try
                {
                    // Add a copy of the shape to the destination sheet at the same position
                    Shape destShape = destSheet.Shapes.AddCopy(
                        srcShape,
                        srcShape.UpperLeftRow,
                        srcShape.UpperLeftColumn,
                        srcShape.LowerRightRow,
                        srcShape.LowerRightColumn);

                    // Preserve size and rotation
                    destShape.Height = srcShape.Height;
                    destShape.Width = srcShape.Width;
                    destShape.RotationAngle = srcShape.RotationAngle;

                    // Preserve alternative text
                    destShape.AlternativeText = srcShape.AlternativeText;

                    // Preserve text and its alignment for text‑based shapes
                    if (!string.IsNullOrEmpty(srcShape.Text))
                    {
                        destShape.Text = srcShape.Text;
                        destShape.TextHorizontalAlignment = srcShape.TextHorizontalAlignment;
                        destShape.TextVerticalAlignment = srcShape.TextVerticalAlignment;
                    }
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Failed to copy shape '{srcShape.Name}': {shapeEx.Message}");
                }
            }

            // Save the workbook with copied shapes
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
