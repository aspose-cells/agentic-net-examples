// Title: Save a merged Aspose.Cells workbook to a MemoryStream for immediate API transmission in C#
// AI Prompts: Generate C# code that merges multiple worksheets with Aspose.Cells and returns the combined workbook as a MemoryStream ready for an ASP.NET Core Web API response. | Provide a method that saves an Aspose.Cells Workbook to a MemoryStream in XLSX format, resets the stream position, and includes robust exception handling. | Show how to stream the merged workbook directly to the client without creating a temporary file on disk.
// Common Searches: how to stream a merged Aspose.Cells workbook from ASP.NET Core Web API | Aspose.Cells C# return merged Excel as MemoryStream for download | save Aspose.Cells workbook to MemoryStream and reset position | export merged workbook to XLSX stream without creating a file C# | C# Aspose.Cells memory stream for HTTP response
// Tags: Aspose.Cells save workbook to memory stream | merged workbook export as XLSX stream | C# ASP.NET Core return Excel MemoryStream | Aspose.Cells workbook merging to stream | reset memory stream position after save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example defines a helper that creates (or merges) an Aspose.Cells Workbook, saves it into a MemoryStream in XLSX format, resets the stream to the beginning, and returns it. A console demo shows how to write the stream to a file, illustrating how the same stream can be sent directly via a web API.
    public static class WorkbookHelper
    {
        /// <returns>MemoryStream containing the workbook data.</returns>
        public static MemoryStream GetMergedWorkbookStream()
        {
            try
            {
                // Initialize a new workbook. Replace with loading an existing workbook if needed.
                Workbook workbook = new Workbook();

                // TODO: Add merging logic here (e.g., copy worksheets, merge data, etc.)

                // Prepare a memory stream to hold the workbook data.
                MemoryStream memoryStream = new MemoryStream();

                // Save the workbook into the memory stream in XLSX format.
                workbook.Save(memoryStream, SaveFormat.Xlsx);

                // Reset the stream position to the beginning for reading.
                memoryStream.Position = 0;

                return memoryStream;
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to let the caller handle it.
                throw new InvalidOperationException("Failed to create or save the merged workbook.", ex);
            }
        }
    }

    // Entry point for the console application.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Get the merged workbook as a memory stream.
                using (MemoryStream stream = WorkbookHelper.GetMergedWorkbookStream())
                {
                    // Define output file path.
                    string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedWorkbook.xlsx");

                    // Write the stream to a file.
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Merged workbook saved successfully to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Log the error and exit.
                Console.Error.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
