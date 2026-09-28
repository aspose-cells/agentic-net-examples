// Title: How to log timestamp and worksheet name each time a WordArt shape is added with Aspose.Cells for .NET
// AI Prompts: Create a C# method that appends a log entry containing the current date‑time and the worksheet name whenever worksheet.Shapes.AddWordArt is invoked. | Update the Aspose.Cells example to write a timestamped record to a text file each time a WordArt object is inserted into a workbook. | Add error‑handled logging around WordArt shape creation in a .NET Excel workbook using Aspose.Cells.
// Common Searches: aspocells c# log when adding wordart shape to worksheet | record timestamp for each shape insertion in Aspose.Cells workbook | write log file for WordArt addition using Aspose.Cells .NET | how to track Excel shape creation events with Aspose.Cells in C#
// Tags: Aspose.Cells WordArt addition logging | timestamp shape creation log .NET | C# write worksheet activity to text file | Excel shape insertion event logging | error-handled logging for Aspose.Cells shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt shape to the first worksheet, and logs the current timestamp along with the worksheet name to a text file each time the shape is added. It includes error handling for shape addition, logging, and workbook saving.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Sheet1";

            try
            {
                // Add a WordArt shape to the worksheet
                // style = 0 (Style1)
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    0,                     // style (0 = Style1)
                    "Sample WordArt",      // text
                    0,                     // upper left row
                    0,                     // upper left column
                    0,                     // top offset (pixels)
                    0,                     // left offset (pixels)
                    50,                    // height (pixels)
                    200                    // width (pixels)
                );

                // Log the addition of the WordArt shape
                LogWordArtAddition(worksheet.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add WordArt: {ex.Message}");
            }

            // Prepare output path
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Logs the timestamp and worksheet name each time a WordArt shape is added
    static void LogWordArtAddition(string worksheetName)
    {
        try
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - WordArt added to worksheet: {worksheetName}";
            // Append the log entry to a text file
            File.AppendAllText("WordArtLog.txt", logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Logging failed: {ex.Message}");
        }
    }
}
