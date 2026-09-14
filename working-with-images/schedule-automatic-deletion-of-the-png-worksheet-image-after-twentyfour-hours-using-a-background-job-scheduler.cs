// Title: Render the first worksheet to a PNG image with Aspose.Cells and automatically delete the file after 24 hours using a background Task in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, exports the first worksheet as a PNG file, and starts a non‑blocking Task that waits 24 hours before deleting the PNG. | Show how to implement a delayed file‑deletion routine using Task.Delay in a console application after creating an image with Aspose.Cells.
// Common Searches: c# Aspose.Cells export first worksheet to png and schedule file removal | how to delete a generated image after a day in a .NET console app | background task for delayed file deletion using Task.Delay | automate cleanup of temporary worksheet images created by Aspose.Cells
// Tags: Aspose.Cells worksheet to PNG export | C# delayed file deletion with Task.Delay | background task cleanup temporary image .NET | schedule automatic removal of generated PNG | export Excel sheet as image and auto‑delete

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageDeletion
{
    // The sample loads an Excel workbook with Aspose.Cells, renders the first worksheet to a PNG file, and launches a background Task that waits 24 hours before deleting the generated image, ensuring non‑blocking cleanup.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Paths to the workbook and the output PNG image
            string workbookPath = @"C:\Data\Sample.xlsx";
            string imagePath = @"C:\Data\WorksheetImage.png";

            try
            {
                // Verify that the workbook file exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(imagePath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    Console.WriteLine("Invalid image path.");
                    return;
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                // Set image export options (default format is PNG)
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                // Render the first worksheet to an image
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], options);
                sheetRender.ToImage(0, imagePath);

                Console.WriteLine($"Worksheet image saved to: {imagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
                return;
            }

            // Schedule automatic deletion of the PNG image after 24 hours
            ScheduleDeletion(imagePath, TimeSpan.FromHours(24));

            // Keep the application running if needed (e.g., console app)
            Console.WriteLine("Image saved and deletion scheduled. Press any key to exit.");
            Console.ReadKey();
        }

        /// <param name="filePath">Full path of the file to delete.</param>
        /// <param name="delay">Time to wait before deletion.</param>
        private static void ScheduleDeletion(string filePath, TimeSpan delay)
        {
            // Run the deletion logic on a background thread without blocking the main thread
            Task.Run(async () =>
            {
                // Wait for the specified delay (e.g., 24 hours)
                await Task.Delay(delay);

                try
                {
                    // Delete the file if it still exists
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                        Console.WriteLine($"File '{filePath}' has been deleted after {delay.TotalHours} hours.");
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors that occur during deletion
                    Console.WriteLine($"Error deleting file '{filePath}': {ex.Message}");
                }
            });
        }
    }
}
