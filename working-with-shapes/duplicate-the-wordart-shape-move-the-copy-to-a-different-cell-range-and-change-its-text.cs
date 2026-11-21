// Title: Copy a WordArt shape to another cell and modify its text using Aspose.Cells for .NET (C#)
// Common Searches: how to duplicate a WordArt shape and move it to cell B5 using Aspose.Cells C# | Aspose.Cells AddCopy method example for copying shapes to a specific row and column | changing the text of a copied WordArt shape in an Excel workbook with Aspose.Cells | C# code to find the first WordArt shape in a worksheet and create a copy | save workbook after updating shape text with Aspose.Cells .NET
// Tags: WordArt shape duplication Aspose.Cells | shape copy API Aspose.Cells | update shape text C# | Excel worksheet shape manipulation .NET | Aspose.Cells shape handling example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample loads an existing Excel file, locates the first WordArt shape on the first worksheet, creates a copy of that shape at cell B5 using the AddCopy method, changes the copied shape's Text property, and saves the workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var sheet = workbook.Worksheets[0];

                // Find the first WordArt shape on the worksheet
                Shape originalWordArt = null;
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsWordArt)
                    {
                        originalWordArt = shape;
                        break;
                    }
                }

                if (originalWordArt != null)
                {
                    // Duplicate the WordArt shape at cell B5 (row 4, column 1)
                    // AddCopy overload requires row, column and pixel offsets
                    Shape copiedWordArt = sheet.Shapes.AddCopy(originalWordArt, 4, 1, 0, 0);

                    // Optionally adjust position (already set via AddCopy)
                    // Change the text of the copied WordArt
                    copiedWordArt.Text = "New WordArt Text";
                }
                else
                {
                    Console.WriteLine("No WordArt shape found in the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
