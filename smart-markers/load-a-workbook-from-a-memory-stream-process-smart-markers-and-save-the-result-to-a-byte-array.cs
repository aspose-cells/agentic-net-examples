// Title: Process Smart Markers in an Excel workbook loaded from a byte array using Aspose.Cells and return the result as a byte array (C#)
// Common Searches: how to use WorkbookDesigner to process smart markers from a byte array in C# | load an XLSX template from memory, apply smart markers and get the output as byte[] with Aspose.Cells | convert processed Excel workbook to byte array without writing to disk using Aspose.Cells | Aspose.Cells smart markers processing with MemoryStream in .NET Core | C# example for reading Excel file into byte[] and saving processed result back to byte[]
// Tags: WorkbookDesigner smart markers memory stream | load workbook from byte array Aspose.Cells | save processed workbook to byte array C# | in‑memory smart markers processing Aspose | exception handling Aspose.Cells smart markers

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example shows how to read an Excel file supplied as a byte array, load it into an Aspose.Cells Workbook via a MemoryStream, process Smart Markers with WorkbookDesigner, and then save the updated workbook back to a MemoryStream in XLSX format, returning the final content as a byte array. A small console program demonstrates reading a template file, invoking the processing method, and writing the resulting byte array to an output file.
    public class SmartMarkerHandler
    {
        /// <param name="inputData">The input workbook data (e.g., an XLSX file) as a byte array.</param>
        /// <returns>The processed workbook data as a byte array.</returns>
        public byte[] ProcessSmartMarkers(byte[] inputData)
        {
            if (inputData == null || inputData.Length == 0)
                throw new ArgumentException("Input data cannot be null or empty.", nameof(inputData));

            try
            {
                // Load the workbook from the input byte array using a memory stream.
                using (var inputStream = new MemoryStream(inputData))
                {
                    var workbook = new Workbook(inputStream);

                    // Process Smart Markers using WorkbookDesigner.
                    var designer = new WorkbookDesigner(workbook);
                    designer.Process();

                    // Save the processed workbook to an output memory stream.
                    using (var outputStream = new MemoryStream())
                    {
                        workbook.Save(outputStream, SaveFormat.Xlsx);
                        return outputStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to preserve stack trace while providing context.
                throw new InvalidOperationException("Failed to process Smart Markers.", ex);
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            const string inputPath = "Template.xlsx";
            const string outputPath = "Result.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Read the input workbook into a byte array.
                byte[] inputData = File.ReadAllBytes(inputPath);

                // Process Smart Markers.
                var handler = new SmartMarkerHandler();
                byte[] resultData = handler.ProcessSmartMarkers(inputData);

                // Write the processed workbook to the output file.
                File.WriteAllBytes(outputPath, resultData);
                Console.WriteLine($"Processing completed. Output saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
