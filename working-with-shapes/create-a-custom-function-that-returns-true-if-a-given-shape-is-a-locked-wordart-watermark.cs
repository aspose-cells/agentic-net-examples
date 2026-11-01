// Title: C# method to check if an Aspose.Cells shape is a locked WordArt watermark
// AI Prompts: Implement a C# function that receives an Aspose.Cells Shape object and returns true only when the shape is a WordArt element and its IsLocked flag is set. | Write code that iterates through all shapes in an Excel workbook using Aspose.Cells and logs the names of shapes that qualify as locked WordArt watermarks. | Create a reusable ShapeHelper class in C# for Aspose.Cells that safely identifies locked WordArt watermark shapes, handling null references and exceptions.
// Common Searches: how to determine if a shape is a locked WordArt watermark using Aspose.Cells in C# | Aspose.Cells C# check shape IsWordArt and IsLocked properties | filter locked WordArt shapes in an Excel file with Aspose.Cells library | C# code to identify WordArt watermarks that are locked in a workbook | detect WordArt watermark shapes in Aspose.Cells workbook programmatically
// Tags: Aspose.Cells shape IsWordArt locked detection | C# locked WordArt watermark identification | Excel workbook shape filtering Aspose.Cells | shape helper utility Aspose.Cells C# | WordArt watermark detection in Excel using Aspose

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example defines a static ShapeHelper.IsLockedWordArtWatermark method that returns true only when the provided Aspose.Cells Shape is non‑null, is a WordArt object (IsWordArt), and has its IsLocked property set, with exception handling that treats any error as a non‑match. The Program class loads a workbook, walks through each worksheet's Shapes collection, calls the helper for each shape, and prints whether the shape is a locked WordArt watermark.
    public static class ShapeHelper
    {
        /// <param name="shape">The shape to evaluate.</param>
        /// <returns>True if the shape is a WordArt object and is locked; otherwise, false.</returns>
        public static bool IsLockedWordArtWatermark(Shape shape)
        {
            try
            {
                if (shape == null)
                    return false;

                if (!shape.IsWordArt)
                    return false;

                if (!shape.IsLocked)
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                // Log the exception and treat the shape as not matching the criteria.
                Console.WriteLine($"Error evaluating shape: {ex.Message}");
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be processed.
            string filePath = "Sample.xlsx";

            // Ensure the input file exists before attempting to load it.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets and their shapes.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        bool isLockedWordArt = ShapeHelper.IsLockedWordArtWatermark(shape);
                        Console.WriteLine($"Shape '{shape.Name}' locked WordArt: {isLockedWordArt}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any errors that occur while processing the workbook.
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
