// Title: Automatically Delete a Worksheet PNG After 24 Hours with System.Threading.Timer in Aspose.Cells for .NET
// Description: Render a worksheet to a PNG using Aspose.Cells, then schedule a System.Threading.Timer to delete the image after 24 hours, ensuring the timer is disposed safely.
// Keywords: Aspose.Cells PNG cleanup | C# timer delete file | System.Threading.Timer Aspose.Cells | schedule file deletion .NET | temporary worksheet image removal | background job cleanup Aspose.Cells | C# delete file after delay | Aspose.Cells image export timer
// Common Searches: how to delete Aspose.Cells generated PNG after a day | C# timer to remove worksheet image after 24 hours | schedule automatic file cleanup for Aspose.Cells exports | using System.Threading.Timer for temporary file deletion in .NET | Aspose.Cells render worksheet to PNG and auto‑purge
// Developer Intent: The developer needs a reliable way to automatically remove a PNG image created from a worksheet after a 24‑hour period.
// Use Cases: Create a short‑lived PNG preview for email attachments and clean it up after one day to prevent storage growth. | Generate on‑demand worksheet screenshots in a web API and schedule their deletion to meet data‑retention policies. | Run batch PNG exports for reporting, then use a timer to purge the files once they are no longer required.
// AI Prompts: Show how to extend the timer logic to delete multiple worksheet PNGs with a configurable interval. | Provide a Hangfire background‑job example that replaces System.Threading.Timer for cleaning up Aspose.Cells images. | Explain best practices for disposing the timer and handling application shutdown when scheduling file deletion.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Render a worksheet to a PNG using Aspose.Cells, then schedule a System.Threading.Timer to delete the image after 24 hours, ensuring the timer is disposed safely.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["A2"].PutValue(12345);

        // Configure image rendering options for PNG
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // Use PNG format
        options.OnePagePerSheet = true;             // Render whole sheet as one page

        // Define the output image path
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string imagePath = Path.Combine(outputDir, "worksheet.png");

        // Render the worksheet to a PNG file using SheetRender (rule)
        SheetRender sheetRender = new SheetRender(worksheet, options);
        sheetRender.ToImage(0, imagePath);
        sheetRender.Dispose();

        // Schedule automatic deletion of the PNG after 24 hours
        Timer deletionTimer = null;
        deletionTimer = new Timer(state =>
        {
            try
            {
                if (File.Exists(imagePath))
                {
                    File.Delete(imagePath);
                    Console.WriteLine($"Deleted image file: {imagePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting image: {ex.Message}");
            }
            finally
            {
                // Dispose the timer after it fires
                deletionTimer?.Dispose();
            }
        }, null, TimeSpan.FromHours(24), Timeout.InfiniteTimeSpan);

        Console.WriteLine($"Image saved to {imagePath}. Deletion scheduled in 24 hours.");
        // Prevent the application from exiting immediately (optional)
        Console.ReadLine();
    }
}
