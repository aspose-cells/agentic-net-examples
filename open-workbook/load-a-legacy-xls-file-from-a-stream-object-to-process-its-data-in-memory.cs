// Title: Load a legacy XLS workbook from a Stream, edit it in memory, and save with Aspose.Cells for .NET
// Description: Demonstrates how to open an old .xls file directly from a Stream using Aspose.Cells, optionally verify the format, reset the stream, read cell values, add rows programmatically, and write the updated workbook to a MemoryStream for further processing or saving.
// Keywords: Aspose.Cells load XLS from stream | detect Excel file format C# | modify workbook in memory | save workbook to MemoryStream | read cell value Aspose.Cells | legacy Excel .xls processing | C# Excel stream handling
// Common Searches: How to open an .xls file from a Stream with Aspose.Cells | Detect Excel file format before loading in C# | Edit Excel workbook in memory using Aspose.Cells | Save Aspose.Cells workbook to a MemoryStream | Read specific cell after loading workbook from Stream
// Developer Intent: Open a legacy .xls file from a Stream, make in‑memory changes, and save the result without creating intermediate files.
// Use Cases: Process uploaded .xls files from an HTTP request stream, add rows, and return the modified file as a response stream. | Batch‑update legacy .xls spreadsheets stored on a network share by reading each via Stream, applying business rules, and writing the updated files back. | Validate an unknown Excel stream, ensure it is a legacy .xls, then load and manipulate it with Aspose.Cells.
// AI Prompts: Generate a C# method that receives a Stream containing an .xls file, uses Aspose.Cells to detect the format, loads it into a Workbook, adds a timestamp column to each row, and returns the modified Workbook. | Write code that loads an .xls workbook from a MemoryStream, reads the value of cell B2, updates cell C3, and then saves the workbook to a byte array.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open an old .xls file directly from a Stream using Aspose.Cells, optionally verify the format, reset the stream, read cell values, add rows programmatically, and write the updated workbook to a MemoryStream for further processing or saving.
public class LegacyXlsLoader
{
    /// <param name="xlsStream">Stream with XLS file data.</param>
    /// <returns>Workbook instance.</returns>
    public static Workbook LoadFromStream(Stream xlsStream)
    {
        if (xlsStream == null)
            throw new ArgumentNullException(nameof(xlsStream));

        // Ensure the stream is at the beginning.
        if (xlsStream.CanSeek)
            xlsStream.Position = 0;

        // Detect format (optional safety check).
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(xlsStream);
        Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");

        // Reset position after detection.
        if (xlsStream.CanSeek)
            xlsStream.Position = 0;

        // Load workbook.
        Workbook workbook = new Workbook(xlsStream);

        // Example: read cell A1.
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

        return workbook;
    }

    /// <summary>
    /// Demonstrates loading, modifying, and saving a legacy XLS file.
    /// </summary>
    public static void RunDemo()
    {
        const string inputPath = "legacyFile.xls";
        const string outputPath = "modifiedLegacyFile.xls";

        // Verify input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Open the legacy file.
            using (FileStream fileStream = File.OpenRead(inputPath))
            {
                Workbook wb = LoadFromStream(fileStream);

                // Add a new row with sample data.
                Worksheet ws = wb.Worksheets[0];
                int lastRow = ws.Cells.MaxDataRow + 1;
                ws.Cells[lastRow, 0].PutValue("New Entry");
                ws.Cells[lastRow, 1].PutValue(DateTime.Now);

                // Save to a memory stream.
                using (MemoryStream outStream = wb.SaveToStream())
                {
                    // Write the memory stream to the output file.
                    using (FileStream outFile = File.Create(outputPath))
                    {
                        outStream.WriteTo(outFile);
                    }
                }

                Console.WriteLine("Workbook processed and saved successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }

    // Entry point required for console application.
    public static void Main(string[] args)
    {
        try
        {
            RunDemo();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
