// Title: Refresh linked pictures in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Shows how to load an .xlsx file, invoke Workbook.UpdateLinkedDataSource to pull the latest external images, and save the workbook, with comprehensive error handling for loading, updating, and saving.
// Keywords: Aspose.Cells | C# | .NET | Refresh linked pictures | UpdateLinkedDataSource | Excel external image links | linked picture refresh | programmatic image update | Workbook.UpdateLinkedDataSource example | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells refresh linked pictures C# | UpdateLinkedDataSource external images .NET | How to refresh picture links in Excel programmatically | Refresh linked images after web server change Aspose | C# code to update linked pictures in workbook
// Developer Intent: Refresh all linked pictures in a workbook after the source images have changed.
// Use Cases: Load an .xlsx containing linked pictures, call UpdateLinkedDataSource to retrieve the newest images from their URLs, and save the refreshed file. | Schedule a nightly job that updates linked graphics in financial or marketing reports before distribution. | Expose a web API that receives updated image URLs, refreshes the linked pictures in the workbook, and returns the modified Excel file.
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, refreshes all linked pictures, handles missing files, and saves the result. | Create a method that enumerates linked picture objects, logs each source URL, calls UpdateLinkedDataSource, and captures any errors. | Provide a complete example that updates external image links, writes the workbook to a new file, and returns a success/failure status message.

using System;
using System.IO;
using Aspose.Cells;

namespace RefreshLinkedPicturesDemo
{
    // Shows how to load an .xlsx file, invoke Workbook.UpdateLinkedDataSource to pull the latest external images, and save the workbook, with comprehensive error handling for loading, updating, and saving.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithLinkedPictures.xlsx";
            const string outputPath = "OutputWithRefreshedPictures.xlsx";

            Workbook workbook = null;

            // Load the workbook only if the file exists.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Refresh linked pictures by updating external data sources.
            try
            {
                workbook.UpdateLinkedDataSource(new Workbook[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while updating linked data sources: {ex.Message}");
            }

            // Save the workbook with refreshed pictures.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
