// Title: Refresh linked pictures in an Excel workbook concurrently with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file using Aspose.Cells, enumerates all Picture objects across worksheets, and refreshes only the linked pictures in parallel with Parallel.ForEach. | Demonstrate how to use reflection to detect the IsLinked property and invoke the Refresh method on Aspose.Cells Picture objects for version‑independent parallel execution.
// Common Searches: how to use Parallel.ForEach to refresh linked images in an Aspose.Cells workbook | c# refresh external pictures in Excel file with Aspose.Cells and reflection | update dozens of linked pictures in Excel efficiently using multithreading | Aspose.Cells picture.IsLinked property check before refreshing | parallel processing of pictures in Excel using Aspose.Cells .NET
// Tags: parallel picture refresh Aspose.Cells | linked image refresh using reflection | Aspose.Cells external picture update | multithreaded Excel picture processing .NET | IsLinked property handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, gathers all Picture objects from every worksheet, and uses Parallel.ForEach together with reflection to identify linked pictures (via the IsLinked property) and invoke their Refresh method. After processing, the workbook is saved, providing a fast, version‑agnostic way to update dozens of external images.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

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

            // Collect all pictures (potentially linked) from all worksheets
            List<Picture> pictures = new List<Picture>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Picture picture in sheet.Pictures)
                {
                    pictures.Add(picture);
                }
            }

            // Refresh each linked picture in parallel (using reflection to stay compatible with different Aspose.Cells versions)
            Parallel.ForEach(pictures, picture =>
            {
                try
                {
                    // Check for an 'IsLinked' property via reflection
                    PropertyInfo isLinkedProp = picture.GetType().GetProperty("IsLinked", BindingFlags.Public | BindingFlags.Instance);
                    bool isLinked = isLinkedProp != null && isLinkedProp.PropertyType == typeof(bool) && (bool)isLinkedProp.GetValue(picture);

                    if (isLinked)
                    {
                        // Invoke the 'Refresh' method via reflection if it exists
                        MethodInfo refreshMethod = picture.GetType().GetMethod("Refresh", BindingFlags.Public | BindingFlags.Instance);
                        refreshMethod?.Invoke(picture, null);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to refresh picture: {ex.Message}");
                }
            });

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
