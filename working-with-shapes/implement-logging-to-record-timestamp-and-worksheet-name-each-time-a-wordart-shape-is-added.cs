// Title: Log Timestamp and Worksheet Name When Adding WordArt Shapes with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds WordArt shapes, and writes a log entry containing the current date‑time and the worksheet name each time a shape is inserted. The log is saved to a text file and the workbook is stored as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | WordArt | shape logging | timestamp | worksheet name | audit trail | Excel automation | log file
// Common Searches: Aspose.Cells log WordArt addition | C# record worksheet name when adding WordArt | timestamp logging for Excel shapes Aspose | how to write shape events to a file with Aspose.Cells | track WordArt creation in .NET workbook
// Developer Intent: Insert WordArt shapes into a worksheet and automatically append a log entry with the current timestamp and the worksheet's name for each insertion.
// Use Cases: Maintain an audit trail of WordArt objects for compliance or review. | Debug shape placement by correlating timestamps with worksheet identifiers. | Generate usage statistics of WordArt across multiple generated workbooks.
// AI Prompts: Create a generic logging method that records timestamp, worksheet name, and shape type for any Aspose.Cells shape. | Show how to switch the log output from a plain text file to a rotating log file or a database connection. | Write unit tests that verify the log contains correct entries after adding WordArt shapes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtLogging
{
    // C# example that creates a workbook, adds WordArt shapes, and writes a log entry containing the current date‑time and the worksheet name each time a shape is inserted. The log is saved to a text file and the workbook is stored as an XLSX file.
    class Program
    {
        // Path to the log file
        private const string LogFilePath = "WordArtLog.txt";

        static void Main()
        {
            // Ensure the log file is empty at start
            File.WriteAllText(LogFilePath, string.Empty);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Add a few WordArt shapes and log each addition
            AddWordArtWithLogging(worksheet, PresetWordArtStyle.WordArtStyle1, "First WordArt", 1, 0, 1, 0, 100, 400);
            AddWordArtWithLogging(worksheet, PresetWordArtStyle.WordArtStyle3, "Second WordArt", 5, 0, 5, 0, 120, 450);
            AddWordArtWithLogging(worksheet, PresetWordArtStyle.WordArtStyle5, "Third WordArt", 10, 0, 10, 0, 150, 500);

            // Save the workbook
            workbook.Save("WordArtWithLogging.xlsx");
        }

        /// <param name="worksheet">Target worksheet.</param>
        /// <param name="style">Preset WordArt style.</param>
        /// <param name="text">Text for the WordArt.</param>
        /// <param name="topRow">Upper left row index.</param>
        /// <param name="top">Vertical offset in pixels.</param>
        /// <param name="leftColumn">Upper left column index.</param>
        /// <param name="left">Horizontal offset in pixels.</param>
        /// <param name="height">Height in pixels.</param>
        /// <param name="width">Width in pixels.</param>
        private static void AddWordArtWithLogging(Worksheet worksheet,
                                                   PresetWordArtStyle style,
                                                   string text,
                                                   int topRow,
                                                   int top,
                                                   int leftColumn,
                                                   int left,
                                                   int height,
                                                   int width)
        {
            // Add the WordArt shape
            Shape wordArt = worksheet.Shapes.AddWordArt(style, text, topRow, top, leftColumn, left, height, width);

            // Prepare log entry with timestamp and worksheet name
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Added WordArt to worksheet '{worksheet.Name}'";

            // Append the log entry to the log file
            File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
        }
    }
}
