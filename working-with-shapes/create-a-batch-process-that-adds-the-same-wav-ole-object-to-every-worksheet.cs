// Title: Batch embed the same WAV audio file as an OLE object into every worksheet using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells Shapes.AddOleObject to insert a WAV file as an OLE object at cell A1 on each worksheet in a workbook. | Create a C# routine that reads a WAV file into a byte array and adds the same OLE object to all sheets, handling missing file errors gracefully. | Save the workbook after batch‑adding the audio OLE shape and ensure the output directory exists.
// Common Searches: aspnet embed same audio OLE object in all Excel sheets with Aspose.Cells | c# batch add wav ole object to each worksheet in a workbook | how to handle missing wav file when inserting OLE objects using Aspose.Cells | add ole object to multiple worksheets Aspose.Cells example
// Tags: shapes.addoleobject embed wav c# | audio ole shape batch insertion Aspose.Cells | multiple worksheets ole object insertion | wav file not found handling Aspose.Cells | excel workbook save with ole objects

using Aspose.Cells;
using System;
using System.IO;

// The example creates (or loads) a workbook, ensures it has at least three worksheets, reads a WAV file into a byte array, and uses Shapes.AddOleObject to place the same audio OLE object at cell A1 of every worksheet. It handles missing audio files and I/O errors, creates the output directory if needed, and saves the workbook as XLSX.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Ensure the workbook has at least three worksheets
            while (workbook.Worksheets.Count < 3)
            {
                workbook.Worksheets.Add($"Sheet{workbook.Worksheets.Count + 1}");
            }

            // Path to the WAV file that will be embedded as an OLE object
            string wavFilePath = @"C:\Temp\sound.wav";

            // Verify the WAV file exists before attempting to embed it
            if (!File.Exists(wavFilePath))
            {
                Console.WriteLine($"Warning: WAV file not found at '{wavFilePath}'. OLE objects will not be added.");
            }
            else
            {
                try
                {
                    // Read the WAV file into a byte array for embedding
                    byte[] oleData = File.ReadAllBytes(wavFilePath);

                    // Add the same WAV OLE object to every worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // AddOleObject(upperLeftRow, upperLeftColumn, height, width, lowerRightRow, lowerRightColumn, oleObjectData)
                        // Placing the object at cell A1 (row 0, column 0) with a size of 100x100 points
                        sheet.Shapes.AddOleObject(0, 0, 100, 100, 0, 0, oleData);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading WAV file: {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputPath = @"C:\Temp\Output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
