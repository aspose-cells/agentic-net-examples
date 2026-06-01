using System;
using System.IO;
using Aspose.Cells;

public static class LegacyXlsProcessor
{
    /// <summary>
    /// Loads a legacy XLS file from the provided stream, processes its data,
    /// and returns the workbook saved into a memory stream.
    /// </summary>
    /// <param name="xlsStream">Stream containing the legacy XLS file.</param>
    /// <returns>MemoryStream with the processed workbook (saved as XLS).</returns>
    public static MemoryStream LoadAndProcessLegacyXls(Stream xlsStream)
    {
        try
        {
            // Ensure the stream supports seeking
            if (!xlsStream.CanSeek)
                throw new ArgumentException("The input stream must support seeking.", nameof(xlsStream));

            // Detect the file format of the incoming stream
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(xlsStream);

            // Reset the stream position after detection so the workbook can read from the start
            xlsStream.Position = 0;

            // Verify that the detected format is the legacy Excel 97‑2003 format
            if (formatInfo.LoadFormat != LoadFormat.Excel97To2003)
                throw new InvalidOperationException($"Expected a legacy XLS format, but detected {formatInfo.LoadFormat}.");

            // Load the workbook from the stream using the constructor that accepts Stream and LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);
            Workbook workbook = new Workbook(xlsStream, loadOptions);

            // ----- Begin processing the workbook -----
            // Example: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cellA1 = sheet.Cells["A1"];
            Console.WriteLine($"Original A1 value: {cellA1.StringValue}");

            // Example modification: append text to A1
            cellA1.PutValue(cellA1.StringValue + " - processed");
            // ----- End processing -----

            // Save the modified workbook back to a memory stream
            MemoryStream resultStream = workbook.SaveToStream();

            // Position the stream at the beginning for the caller
            resultStream.Position = 0;
            return resultStream;
        }
        catch (Exception)
        {
            // Rethrow to allow caller to handle or log
            throw;
        }
    }
}

public static class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Prevent FileNotFoundException for the input file
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Open the input file as a read‑only stream
            using (FileStream inputStream = File.OpenRead(inputPath))
            {
                // Process the legacy XLS and obtain the result stream
                MemoryStream processedStream = LegacyXlsProcessor.LoadAndProcessLegacyXls(inputStream);

                // Write the processed workbook to the output file
                using (FileStream outputStream = File.Create(outputPath))
                {
                    processedStream.CopyTo(outputStream);
                }

                Console.WriteLine($"Processed file saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}