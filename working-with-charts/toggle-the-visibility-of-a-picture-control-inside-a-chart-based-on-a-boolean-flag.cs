// Title: Toggle picture visibility in an Excel worksheet or chart using a Boolean flag with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide worksheet pictures based on a boolean variable | How to programmatically remove chart overlay images using Aspose.Cells .NET | Aspose.Cells example for deleting all pictures from a workbook when a condition is false | Toggle picture visibility in an Excel file with Aspose.Cells C# code sample | Remove pictures from an Excel worksheet using the Aspose.Cells API in .NET
// Tags: conditional picture removal Aspose.Cells | worksheet pictures delete C# | chart overlay image hide Aspose.Cells | boolean flag picture visibility .NET | Aspose.Cells picture management example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The C# program loads an Excel workbook, accesses the first worksheet and its first chart, and uses a Boolean flag (showPicture) to decide whether to keep or delete all pictures on the worksheet. When the flag is false, it iterates backward through the worksheet's Pictures collection, safely removes each picture, and then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the first chart in the worksheet (adjust index if needed)
                Chart chart = worksheet.Charts[0];

                // Flag that determines picture visibility
                bool showPicture = true; // Set to false to hide the picture

                // If pictures should be hidden, remove them from the worksheet
                if (!showPicture)
                {
                    // Remove all pictures from the worksheet safely by iterating backwards
                    for (int i = worksheet.Pictures.Count - 1; i >= 0; i--)
                    {
                        try
                        {
                            worksheet.Pictures.RemoveAt(i);
                        }
                        catch (Exception picEx)
                        {
                            Console.WriteLine($"Failed to delete picture at index {i}: {picEx.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
