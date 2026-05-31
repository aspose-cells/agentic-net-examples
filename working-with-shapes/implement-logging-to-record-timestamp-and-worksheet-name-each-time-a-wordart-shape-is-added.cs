using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtLogging
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Path for the log file
            string logPath = "WordArtLog.txt";

            // Ensure the log file is empty at start
            File.WriteAllText(logPath, string.Empty);

            // Helper method to add WordArt and log the action
            void AddWordArtWithLog(PresetWordArtStyle style, string text,
                                   int topRow, int top, int leftColumn, int left,
                                   int height, int width)
            {
                // Add the WordArt shape
                Shape wordArt = worksheet.Shapes.AddWordArt(style, text,
                                                            topRow, top, leftColumn, left,
                                                            height, width);

                // Prepare log entry with timestamp and worksheet name
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Worksheet: {worksheet.Name} - Added WordArt: \"{text}\"";

                // Append the log entry to the file
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }

            // Example: add two WordArt shapes and log each addition
            AddWordArtWithLog(PresetWordArtStyle.WordArtStyle1, "First WordArt",
                              2, 0, 2, 0, 100, 300);

            AddWordArtWithLog(PresetWordArtStyle.WordArtStyle3, "Second WordArt",
                              5, 0, 5, 0, 120, 350);

            // Save the workbook
            workbook.Save("WordArtWithLogging.xlsx");
        }
    }
}