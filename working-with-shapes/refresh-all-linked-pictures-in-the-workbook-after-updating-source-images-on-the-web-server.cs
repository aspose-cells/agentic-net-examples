// Title: Enumerate pictures in every worksheet and handle the lack of linked image refresh support using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens an existing .xlsx file, loops through all worksheets, and prints each picture's name and its sheet location. | Show how to attempt a refresh of external linked images in an Aspose.Cells workbook with C#, including error handling that reports the API limitation when refresh is unavailable.
// Common Searches: c# aspocells list all pictures in workbook worksheets | aspocells refresh linked images not supported | how to get picture names from each sheet using aspocells .net | enumerate picture objects in excel file with aspocells c# | aspocells external image link update after web server change
// Tags: Aspose.Cells picture enumeration C# | linked image refresh limitation Aspose.Cells | iterate worksheets pictures Aspose.Cells | log picture metadata Aspose.Cells | external image link handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file with Aspose.Cells, iterates through each worksheet's picture collection, logs each picture's name, notes that the API does not provide a method to refresh linked images, ensures the output directory exists, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all pictures in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Picture picture in sheet.Pictures)
                    {
                        try
                        {
                            // Log picture information; linked picture refresh is not supported via API
                            Console.WriteLine($"Picture '{picture.Name}' found on sheet '{sheet.Name}'.");
                        }
                        catch (Exception picEx)
                        {
                            Console.Error.WriteLine($"Failed to process picture on sheet '{sheet.Name}': {picEx.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook after processing
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
