// Title: Copy all pictures from the first worksheet to a new worksheet while preserving original dimensions using Aspose.Cells for .NET
// AI Prompts: Iterate over the Pictures collection of the first worksheet and invoke the Picture.CopyTo method (using reflection if necessary) to duplicate each image onto a newly added worksheet, keeping its original size. | Create a worksheet named "CopiedPictures", copy every picture from the source sheet to this sheet without altering dimensions, and save the workbook.
// Common Searches: Aspose.Cells C# copy pictures from one sheet to another preserving size | How to duplicate images in Excel using Aspose.Cells .NET | Copy all shapes from first worksheet to new worksheet Aspose.Cells | Use reflection to call Picture.CopyTo in Aspose.Cells when method not directly accessible | Copy pictures between worksheets without losing dimensions Aspose.Cells
// Tags: Aspose.Cells transfer pictures across worksheets | preserve image size when copying Excel shapes | invoke CopyTo on picture via reflection | add new worksheet and transfer images C# | duplicate Excel pictures Aspose.Cells example

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Reflection;

// The example loads an existing workbook, adds a worksheet called "CopiedPictures", loops through each Picture in the first worksheet, and copies the picture to the new sheet using the Picture.CopyTo method accessed via reflection (if available) so the images retain their original dimensions. The modified workbook is then saved as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Reference the first worksheet (source)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet to receive the copied pictures
            Worksheet targetSheet = workbook.Worksheets.Add("CopiedPictures");

            // Copy each picture from source to target
            foreach (Picture picture in sourceSheet.Pictures)
            {
                try
                {
                    // Attempt to use the built‑in CopyTo method via reflection (available in newer versions)
                    MethodInfo copyToMethod = picture.GetType().GetMethod("CopyTo", new[] { typeof(Worksheet) });
                    if (copyToMethod != null)
                    {
                        copyToMethod.Invoke(picture, new object[] { targetSheet });
                    }
                    else
                    {
                        // Fallback: skip copying if CopyTo is unavailable
                        Console.WriteLine($"CopyTo method not available for picture '{picture.Name}'. Skipping.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy picture '{picture.Name}': {ex.Message}");
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
