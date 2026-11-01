// Title: How to log worksheet name and timestamp each time a WordArt shape is added with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells log each AddWordArt call with worksheet name and datetime | record shape addition events to a text file using Aspose.Cells .NET | append timestamped log entries when inserting WordArt into Excel workbook in C# | track worksheet name for every WordArt shape created with Aspose.Cells | write custom log for WordArt shape creation Aspose.Cells C# example
// Tags: AddWordArt logging Aspose.Cells | timestamped shape addition .NET | worksheet name log file C# | Aspose.Cells shape event tracking | Excel WordArt audit log

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a WordArt shape to a worksheet, and appends a log entry containing the current date‑time and the worksheet name to a text file each time the shape is added, then saves the workbook.
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

            // Add a WordArt shape and log the action
            AddWordArt(worksheet, "Sample WordArt");

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Adds a WordArt shape to the specified worksheet and logs the timestamp and worksheet name
    static void AddWordArt(Worksheet worksheet, string text)
    {
        try
        {
            // Add WordArt shape (type 0 = WordArt style 1)
            // Signature: AddWordArt(int shapeType, string text, int upperLeftRow, int upperLeftColumn,
            //                         int top, int left, int height, int width)
            worksheet.Shapes.AddWordArt(0, text, 0, 0, 0, 0, 200, 50);

            // Prepare log entry with current timestamp and worksheet name
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\tWorksheet: {worksheet.Name}";

            // Append the log entry to a log file
            string logPath = "WordArtLog.txt";
            File.AppendAllText(logPath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to add WordArt: {ex.Message}");
        }
    }
}
