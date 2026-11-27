// Title: Log timestamp and worksheet name each time a WordArt shape is added using Aspose.Cells for .NET
// Common Searches: c# log worksheet name when adding WordArt with Aspose.Cells | record timestamp for each WordArt shape created in an Excel workbook using Aspose.Cells | Aspose.Cells .NET write custom log file for shape additions | how to verify IsWordArt property while logging shape creation in Aspose.Cells | save WordArt shape details to a text file with Aspose.Cells for .NET
// Tags: log WordArt shape addition Aspose.Cells | timestamp logging for worksheet shapes .NET | custom text file logging of Aspose.Cells shape creation | verify IsWordArt property during shape insertion | record worksheet name with shape logging Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtLoggingDemo
{
    // Demonstrates adding WordArt shapes to multiple worksheets and writing an ISO‑8601 timestamp, worksheet name, shape name, and IsWordArt flag to a custom text log using Aspose.Cells for .NET.
    class Program
    {
        // Path to the log file
        private const string LogFilePath = "WordArtLog.txt";

        static void Main(string[] args)
        {
            try
            {
                // Ensure any previous log is cleared
                if (File.Exists(LogFilePath))
                {
                    File.Delete(LogFilePath);
                }

                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Add a few worksheets for demonstration
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "SalesData";

                Worksheet sheet2 = workbook.Worksheets.Add("Summary");
                Worksheet sheet3 = workbook.Worksheets.Add("Archive");

                // Add WordArt to each worksheet and log the operation
                AddWordArtWithLogging(sheet1, PresetWordArtStyle.WordArtStyle2, "Quarterly Sales", 2, 0, 2, 0, 100, 400);
                AddWordArtWithLogging(sheet2, PresetWordArtStyle.WordArtStyle5, "Yearly Summary", 1, 0, 1, 0, 120, 500);
                AddWordArtWithLogging(sheet3, PresetWordArtStyle.WordArtStyle9, "Archived Data", 3, 0, 3, 0, 80, 300);

                // Save the workbook to verify the shapes are persisted
                string outputPath = "WordArtDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine($"Logging completed. See '{LogFilePath}' for details.");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Log($"ERROR: {ex.Message}");
                Console.WriteLine("An error occurred. Check the log for details.");
            }
        }

        /// <param name="worksheet">Target worksheet.</param>
        /// <param name="style">Preset WordArt style.</param>
        /// <param name="text">Text to display in the WordArt.</param>
        /// <param name="topRow">Upper left row index.</param>
        /// <param name="top">Vertical offset in pixels.</param>
        /// <param name="leftColumn">Upper left column index.</param>
        /// <param name="left">Horizontal offset in pixels.</param>
        /// <param name="height">Height in pixels.</param>
        /// <param name="width">Width in pixels.</param>
        private static void AddWordArtWithLogging(
            Worksheet worksheet,
            PresetWordArtStyle style,
            string text,
            int topRow,
            int top,
            int leftColumn,
            int left,
            int height,
            int width)
        {
            // Add the WordArt shape using the exact SDK signature
            Shape wordArt = worksheet.Shapes.AddWordArt(
                style,
                text,
                topRow,
                top,
                leftColumn,
                left,
                height,
                width);

            // Verify that the shape is indeed WordArt (property IsWordArt)
            bool isWordArt = wordArt.IsWordArt;

            // Build the log entry
            string timestamp = DateTime.Now.ToString("o"); // ISO 8601 format
            string logEntry = $"{timestamp} | Worksheet: '{worksheet.Name}' | Shape: '{wordArt.Name}' | IsWordArt: {isWordArt}";

            // Write to log
            Log(logEntry);
        }

        /// <summary>
        /// Appends a line to the log file and also writes it to the console.
        /// </summary>
        /// <param name="message">Message to log.</param>
        private static void Log(string message)
        {
            try
            {
                // Append to the log file
                File.AppendAllText(LogFilePath, message + Environment.NewLine);
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"Failed to write to log file: {ioEx.Message}");
            }

            // Also output to console for immediate feedback
            Console.WriteLine(message);
        }
    }
}
