// Title: Log Timestamp and Worksheet Name When Adding WordArt Shapes with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add WordArt shapes, and automatically record each addition with a timestamp and the worksheet name in a text log before saving the file.
// Keywords: Aspose.Cells | C# | .NET | WordArt | shape logging | timestamp | worksheet name | Excel automation | AddWordArt | log to file | audit trail
// Common Searches: Aspose.Cells log WordArt creation | C# record worksheet name when adding shape | timestamp logging for WordArt in .NET | write shape addition details to file Aspose.Cells | how to audit WordArt objects in Excel using Aspose
// Developer Intent: Add WordArt shapes and automatically write a log entry with the current date‑time and worksheet name for each shape.
// Use Cases: Maintain an audit trail of WordArt objects for compliance reporting. | Debug shape placement by reviewing a chronological log of additions across worksheets. | Export a summary of WordArt content and locations for downstream processing.
// AI Prompts: Generate C# code that logs any Aspose.Cells shape (e.g., pictures, charts) with timestamp, worksheet name, and shape type. | Show how to extend the logging method to include cell coordinates and output the log in CSV format. | Explain how to implement thread‑safe logging when adding WordArt shapes concurrently in a multi‑threaded Aspose.Cells application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtLogging
{
    // Demonstrates how to create a workbook, add WordArt shapes, and automatically record each addition with a timestamp and the worksheet name in a text log before saving the file.
    class Program
    {
        // Path to the log file
        private const string LogFilePath = "WordArtLog.txt";

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Ensure the log file is empty at start
            File.WriteAllText(LogFilePath, string.Empty);

            // Add a WordArt shape and log the operation
            AddWordArtWithLogging(
                worksheet,
                PresetWordArtStyle.WordArtStyle2,
                "Hello Aspose!",
                topRow: 2,
                top: 0,
                leftColumn: 2,
                left: 0,
                height: 100,
                width: 400);

            // Add another WordArt shape to demonstrate multiple logs
            AddWordArtWithLogging(
                worksheet,
                PresetWordArtStyle.WordArtStyle5,
                "Second WordArt",
                topRow: 5,
                top: 0,
                leftColumn: 5,
                left: 0,
                height: 120,
                width: 450);

            // Save the workbook
            workbook.Save("WordArtWithLogging.xlsx");
        }

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
            // Add the WordArt shape
            Shape wordArt = worksheet.Shapes.AddWordArt(
                style,
                text,
                topRow,
                top,
                leftColumn,
                left,
                height,
                width);

            // Prepare log entry with timestamp and worksheet name
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\tWorksheet: {worksheet.Name}\tShape: {text}";

            // Append the log entry to the log file
            File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
        }
    }
}
