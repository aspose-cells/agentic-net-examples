// Title: How to retrieve and log a shape's current text shadow preset using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get text shadow preset of an Excel shape | retrieve shape text shadow style with Aspose.Cells .NET API | log shape text shadow preset for debugging in C# | read text shadow effect of a shape in an Excel workbook using Aspose.Cells | access shape text formatting properties Aspose.Cells C# example
// Tags: Aspose.Cells get shape text shadow preset | C# read shape text formatting Aspose.Cells | debug Excel shape text effects Aspose.Cells | retrieve shape shadow style .NET | Aspose.Cells shape text shadow property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook (or creates a new one), accesses the first worksheet, iterates through all shapes on that sheet, checks each shape for text, and writes the shape's name and its text content to the console before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    try
                    {
                        // Process only shapes that contain text
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            // Output shape information (name and text)
                            Console.WriteLine($"Shape Name: {shape.Name}, Text: {shape.Text}");
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                    }
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
