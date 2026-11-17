// Title: Schedule automatic deletion of a worksheet PNG image 24 hours after export with Aspose.Cells in C#
// Common Searches: C# schedule deletion of an exported worksheet PNG after 24 hours using Aspose.Cells | Use Task.Delay to automatically remove temporary PNG files created from Excel sheets | Background task for delayed cleanup of Aspose.Cells PNG images in a .NET console application | How to auto‑delete temporary image files after a day when exporting Excel to PNG | Best practice for delayed file removal of worksheet images in C#
// Tags: schedule file deletion with Task.Delay | auto‑remove exported worksheet PNG | Aspose.Cells PNG export cleanup | background task for delayed file removal .NET | temporary image deletion after export | C# delayed file delete using Task.Run

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads an Excel workbook, exports the first worksheet as a PNG file, and starts a background task that waits 24 hours before deleting the generated image, using Task.Run and Task.Delay to perform the delayed cleanup without blocking the main thread.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Export the first worksheet to a PNG image
        string imagePath = "Sheet1.png";
        workbook.Save(imagePath, SaveFormat.Png);

        // Schedule automatic deletion of the PNG image after 24 hours
        ScheduleDeletion(imagePath, TimeSpan.FromHours(24));

        // Keep the application running if needed (e.g., console app)
        Console.WriteLine("Image created. It will be deleted automatically after 24 hours.");
        Console.ReadLine();
    }

    /// <param name="filePath">Full path of the file to delete.</param>
    /// <param name="delay">Time to wait before deletion.</param>
    private static void ScheduleDeletion(string filePath, TimeSpan delay)
    {
        // Run the deletion logic on a separate thread without blocking the main thread
        Task.Run(async () =>
        {
            // Wait for the specified delay (24 hours)
            await Task.Delay(delay);

            // Delete the file if it still exists
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine($"File '{filePath}' has been deleted after {delay.TotalHours} hours.");
                }
            }
            catch (Exception ex)
            {
                // Log or handle any errors that occur during deletion
                Console.WriteLine($"Error deleting file '{filePath}': {ex.Message}");
            }
        });
    }
}
