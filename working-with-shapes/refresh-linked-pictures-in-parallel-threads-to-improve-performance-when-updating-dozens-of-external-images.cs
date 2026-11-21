// Title: Using Parallel.ForEach to Process Linked Pictures in an Excel Workbook with Aspose.Cells for .NET
// Common Searches: how to iterate over all pictures in an Aspose.Cells workbook using C# | parallel processing of linked images in Excel with Aspose.Cells .NET | update external picture objects in multiple worksheets using Aspose.Cells | C# example for collecting and handling picture objects in Aspose.Cells | multithreaded picture handling in Excel files with Aspose.Cells library
// Tags: parallel picture processing Aspose.Cells | collect worksheet pictures C# | linked image handling Excel Aspose.Cells | multithreaded Aspose.Cells picture iteration | save workbook after picture update Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, gathers every Picture object from all worksheets into a list, and attempts to refresh each picture inside a Parallel.ForEach loop (the Refresh method is not available in the current API). After the parallel iteration, the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Collect all pictures from every worksheet (linked picture detection not available in this version)
            List<Picture> pictures = new List<Picture>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Picture pic in sheet.Pictures)
                {
                    pictures.Add(pic);
                }
            }

            // Attempt to refresh each picture (Refresh method not available in this version, so this step is skipped)
            Parallel.ForEach(pictures, pic =>
            {
                try
                {
                    // If a Refresh method becomes available in future versions, it can be called here.
                    // Example: pic.Refresh();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process picture: {ex.Message}");
                }
            });

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
