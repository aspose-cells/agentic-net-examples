// Title: Merge Multiple Excel Files to a MemoryStream with Aspose.Cells for .NET
// Description: C# code that validates an array of Excel paths, loads each workbook with Aspose.Cells, combines them using the Combine method, saves the result to a MemoryStream in XLSX format, resets the stream position, and returns it for immediate use in a web API or other downstream services.
// Keywords: Aspose.Cells merge workbooks C# | combine Excel files memory stream | Aspose.Cells SaveFormat.Xlsx | return merged workbook as stream | .NET Excel merge API | in‑memory Excel consolidation
// Common Searches: how to merge several Excel files into one workbook using Aspose.Cells | Aspose.Cells combine workbooks and get MemoryStream | C# save merged Excel workbook to stream for HTTP response | Aspose.Cells merge files without writing to disk | memory stream output of combined Excel sheets .NET
// Developer Intent: Provide a reusable method that merges multiple Excel files into a single workbook and returns the result as a MemoryStream for immediate transmission.
// Use Cases: Expose an ASP.NET Core endpoint that receives uploaded Excel files, merges them with ExcelMerger.MergeFiles, and streams the combined XLSX back to the client. | Create a scheduled service that consolidates departmental spreadsheets, stores the MemoryStream in cloud blob storage, and avoids temporary files on the server. | Implement a microservice that aggregates reporting data from several sources and returns the merged workbook directly over HTTP.
// AI Prompts: Generate an ASP.NET Core controller action that calls ExcelMerger.MergeFiles and returns the MemoryStream as a FileResult with the correct XLSX content type. | Add structured logging to the MergeFiles method and throw a custom MergeException that includes the problematic file path. | Refactor MergeFiles to accept an IEnumerable<IFormFile>, merge the uploads in memory, and output the combined workbook as a stream.

using System;
using System.IO;
using Aspose.Cells;

namespace MyApi
{
    // C# code that validates an array of Excel paths, loads each workbook with Aspose.Cells, combines them using the Combine method, saves the result to a MemoryStream in XLSX format, resets the stream position, and returns it for immediate use in a web API or other downstream services.
    public static class ExcelMerger
    {
        /// <param name="filePaths">Array of full file paths to the source Excel files.</param>
        /// <returns>MemoryStream containing the merged workbook.</returns>
        /// <exception cref="ArgumentException">Thrown when no file paths are provided.</exception>
        /// <exception cref="FileNotFoundException">Thrown when any of the specified files do not exist.</exception>
        /// <exception cref="InvalidOperationException">Thrown when a file cannot be processed.</exception>
        public static MemoryStream MergeFiles(string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0)
                throw new ArgumentException("No file paths provided.", nameof(filePaths));

            // Destination workbook that will hold the merged content.
            var mergedWorkbook = new Workbook();

            foreach (var path in filePaths)
            {
                // Ensure the source file exists before attempting to load it.
                if (!File.Exists(path))
                    throw new FileNotFoundException($"File not found: {path}");

                try
                {
                    // Load the source workbook.
                    var sourceWorkbook = new Workbook(path);

                    // Merge the source workbook into the destination workbook.
                    mergedWorkbook.Combine(sourceWorkbook);
                }
                catch (Exception ex)
                {
                    // Wrap any exception with context about the file that caused it.
                    throw new InvalidOperationException($"Failed to process file '{path}'.", ex);
                }
            }

            // Save the merged workbook to a memory stream in XLSX format.
            var stream = new MemoryStream();
            mergedWorkbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0; // Reset position for downstream consumers.

            return stream;
        }
    }

    internal class Program
    {
        // Entry point for the console application.
        private static void Main(string[] args)
        {
            // Expected usage: dotnet run <outputFilePath> <inputFilePath1> <inputFilePath2> [...]
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <outputFilePath> <inputFilePath1> [<inputFilePath2> ...]");
                return;
            }

            string outputPath = args[0];
            string[] inputPaths = new string[args.Length - 1];
            Array.Copy(args, 1, inputPaths, 0, inputPaths.Length);

            try
            {
                using (MemoryStream mergedStream = ExcelMerger.MergeFiles(inputPaths))
                {
                    // Ensure the output directory exists.
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Write the merged workbook to the specified output file.
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        mergedStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Merged workbook saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
