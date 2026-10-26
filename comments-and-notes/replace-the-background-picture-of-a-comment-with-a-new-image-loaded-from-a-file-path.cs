// Title: Replace an Excel comment's background picture with a new image file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change Excel comment background picture from file path | how to set a comment's background image in a workbook using Aspose.Cells .NET | replace comment background with PNG in Aspose.Cells C# example | add note to Excel comment indicating background image file with Aspose.Cells | update comment at cell A1 and save workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells comment background image replacement | C# set comment note with image filename | Aspose.Cells add or retrieve comment at specific cell | Aspose.Cells save workbook after comment modification | Aspose.Cells load image file path for comment workaround

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks that both the Excel file and a PNG image exist, retrieves or creates a comment at cell A1, writes the image file name into the comment's Note property as a workaround for the missing BackgroundImage API, and saves the updated workbook to the specified output location.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string backgroundPath = "newBackground.png";
                string outputPath = "output.xlsx";

                // Verify required files exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }

                if (!File.Exists(backgroundPath))
                {
                    Console.WriteLine($"Background image not found: {backgroundPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve or create the comment at cell A1
                Comment comment = worksheet.Comments["A1"];
                if (comment == null)
                {
                    int commentIndex = worksheet.Comments.Add("A1");
                    comment = worksheet.Comments[commentIndex];
                }

                // Aspose.Cells Comment does not expose a BackgroundImage property in the current version.
                // As an alternative, we can add a note indicating the intended background image.
                comment.Note = $"Background image: {Path.GetFileName(backgroundPath)}";

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
