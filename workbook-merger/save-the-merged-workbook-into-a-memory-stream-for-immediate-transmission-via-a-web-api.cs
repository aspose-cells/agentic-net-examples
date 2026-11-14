// Title: Convert a merged Aspose.Cells Workbook to an XLSX MemoryStream for immediate API response in C#
// Common Searches: how to return an Excel file from an ASP.NET Core Web API using Aspose.Cells | save Aspose.Cells workbook to MemoryStream without writing to disk C# | stream merged workbook as XLSX in a .NET web service with Aspose.Cells | Aspose.Cells convert workbook to byte array for HTTP response | C# Aspose.Cells export merged workbook to memory for API delivery
// Tags: Aspose.Cells save workbook to MemoryStream | export merged workbook as XLSX in C# | ASP.NET Core return Excel file from memory | convert workbook to byte array Aspose.Cells | exception handling Aspose.Cells workbook save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example provides a helper method that validates a non‑null Workbook, creates a MemoryStream, saves the workbook into the stream using the XLSX format, resets the stream position, and returns the stream while wrapping any errors in an InvalidOperationException. The demo program creates a workbook, adds data, converts it to a MemoryStream via the helper, and writes the stream to a file, illustrating how to prepare a merged workbook for immediate transmission through a web API.
    public static class WorkbookHelper
    {
        /// <param name="workbook">The merged Aspose.Cells Workbook instance.</param>
        /// <returns>A MemoryStream containing the workbook data.</returns>
        public static MemoryStream SaveWorkbookToStream(Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var memoryStream = new MemoryStream();
            try
            {
                // Save the workbook into the memory stream using the XLSX format.
                workbook.Save(memoryStream, SaveFormat.Xlsx);

                // Reset the stream position to the beginning for reading.
                memoryStream.Position = 0;

                return memoryStream;
            }
            catch (Exception ex)
            {
                // Wrap any exception for clearer diagnostics.
                throw new InvalidOperationException("Failed to save workbook to stream.", ex);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add sample data.
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Hello Aspose.Cells");

                // Convert workbook to a memory stream.
                using var stream = WorkbookHelper.SaveWorkbookToStream(workbook);

                // Define output file path.
                string outputPath = "output.xlsx";

                // Write the stream content to the file system.
                File.WriteAllBytes(outputPath, stream.ToArray());

                Console.WriteLine($"Workbook successfully saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
