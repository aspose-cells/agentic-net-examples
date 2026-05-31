using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class AutoFillSeriesFromStream
    {
        /// <summary>
        /// Loads a workbook from the provided stream, creates a numeric series in column A,
        /// and autofills the series down to the target range.
        /// </summary>
        /// <param name="inputStream">Stream containing the source Excel file.</param>
        /// <param name="outputPath">File path where the modified workbook will be saved.</param>
        public static void Run(Stream inputStream, string outputPath)
        {
            try
            {
                // Ensure the stream is positioned at the beginning before loading.
                if (inputStream.CanSeek)
                {
                    inputStream.Position = 0;
                }

                // Load the workbook from the stream.
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Initialize the source range with the first two numbers of the series.
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);

                // Create the source range (the pattern to be extended).
                AsposeRange sourceRange = cells.CreateRange("A1:A2");

                // Define the target range where the series will be filled (A3:A10).
                AsposeRange targetRange = cells.CreateRange("A3:A10");

                // Perform the autofill using the Series type to extend the numeric pattern.
                sourceRange.AutoFill(targetRange, AutoFillType.Series);

                // Ensure the output directory exists.
                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the modified workbook to the specified output file.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
                throw;
            }
        }

        // Example usage.
        public static void Main()
        {
            try
            {
                // Create a sample workbook in memory to demonstrate the method.
                using (MemoryStream sampleStream = new MemoryStream())
                {
                    // Create a simple workbook and save it to the memory stream.
                    Workbook sampleWorkbook = new Workbook();
                    sampleWorkbook.Worksheets[0].Name = "Data";
                    sampleWorkbook.Save(sampleStream, SaveFormat.Xlsx);

                    // Call the autofill method, passing the stream and desired output path.
                    Run(sampleStream, "AutoFilledSeries.xlsx");
                }

                Console.WriteLine("Workbook processed and saved as AutoFilledSeries.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}