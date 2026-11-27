// Title: How to assign different background images to worksheets in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set worksheet background image from file path | apply rotating background pictures to multiple Excel sheets with Aspose.Cells | C# loop through worksheets and assign background images cyclically Aspose.Cells | load PNG or JPG image bytes into worksheet background Aspose.Cells .NET | save workbook with individual sheet backgrounds using Aspose.Cells
// Tags: assign background picture to Excel worksheet Aspose.Cells | rotate image list across multiple sheets C# | read image file into byte array for sheet background | export workbook with individual sheet backgrounds | index-driven worksheet background selection

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    // The example creates a new Workbook, adds several worksheets, defines a list of image file paths, and iterates over each worksheet. For each sheet it selects an image cyclically, reads the file into a byte array, and sets the worksheet's BackgroundImage property. It validates file existence, creates the output folder if needed, and saves the workbook as an XLSX file with distinct backgrounds per sheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: use create)
                Workbook workbook = new Workbook();

                // Add sample worksheets for demonstration
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");
                workbook.Worksheets.Add("Sheet4");

                // Predefined collection of background image file paths
                // Ensure these files exist at the specified locations
                List<string> backgroundImages = new List<string>
                {
                    @"C:\Images\Background1.png",
                    @"C:\Images\Background2.jpg",
                    @"C:\Images\Background3.bmp"
                };

                // Apply background images to each worksheet based on its index
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Select image cyclically from the collection
                    string imagePath = backgroundImages[i % backgroundImages.Count];

                    // Verify that the image file exists before loading
                    if (File.Exists(imagePath))
                    {
                        // Load image bytes and assign to the worksheet background
                        byte[] imageBytes = File.ReadAllBytes(imagePath);
                        sheet.BackgroundImage = imageBytes;
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Background image not found at '{imagePath}'. Skipping assignment for sheet '{sheet.Name}'.");
                    }
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Output\WorkbookWithBackgrounds.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file (lifecycle rule: use save)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
