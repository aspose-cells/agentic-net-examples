// Title: Read a shape's width and height and write the dimensions to Excel cells with Aspose.Cells for .NET
// AI Prompts: Load "input.xlsx" if it exists; otherwise create a new Workbook, locate the shape named "MyShape" on the first worksheet, read its Width and Height properties, and store the numeric values in cells A2 and B2. | Write the labels "Width" and "Height" to cells A1 and B1, then save the workbook as "output.xlsx" after extracting the shape's bounding box dimensions.
// Common Searches: Aspose.Cells C# get shape width and height from worksheet | how to write shape dimensions to specific Excel cells using Aspose.Cells | read shape bounding box measurements with Aspose.Cells .NET | create workbook when missing and extract shape size Aspose.Cells example | Aspose.Cells example reading shape size and saving to new file
// Tags: Aspose.Cells read shape dimensions | C# write shape size to worksheet cells | shape bounding box extraction Aspose | missing workbook handling Aspose.Cells | retrieve shape width height .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeBoundingBoxExample
{
    // The example loads an existing workbook (or creates a new one if the file is absent), accesses the first worksheet, finds a shape named "MyShape", reads its Width and Height properties, writes the labels and values to cells A1:B2, and saves the result to an output Excel file.
    class Program
    {
        static void Main(string[] args)
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
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape by its name.
                Shape shape = worksheet.Shapes["MyShape"]; // Change "MyShape" to your shape's name.

                if (shape == null)
                {
                    Console.WriteLine("Shape \"MyShape\" not found in the worksheet.");
                    return;
                }

                // Use the shape's Width and Height properties to obtain its dimensions.
                double width = shape.Width;
                double height = shape.Height;

                // Write the dimensions to worksheet cells.
                worksheet.Cells["A1"].PutValue("Width");
                worksheet.Cells["B1"].PutValue("Height");
                worksheet.Cells["A2"].PutValue(width);
                worksheet.Cells["B2"].PutValue(height);

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
