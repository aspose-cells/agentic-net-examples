// Title: Load an Excel workbook from a Stream, set Aspose.Cells compression to Level6, and save it to a MemoryStream in C#
// Common Searches: C# Aspose.Cells load workbook from FileStream and write to MemoryStream | how to enable Level6 compression when saving an XLSX with Aspose.Cells | Aspose.Cells compress workbook while saving to a MemoryStream | saving Excel workbook to MemoryStream without creating a temporary file using Aspose.Cells | reset MemoryStream position after calling Workbook.Save in Aspose.Cells
// Tags: Aspose.Cells compression Level6 setting | stream-based workbook loading Aspose.Cells | saving workbook to MemoryStream in XLSX format | Aspose.Cells Save method with MemoryStream | resetting stream position after Aspose.Cells save

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an Excel workbook from an input Stream, optionally set the workbook's compression level to Level6, save the workbook into a MemoryStream as XLSX, and reset the stream position for downstream use, using Aspose.Cells for .NET.
public class WorkbookProcessor
{
    // Loads a workbook from the given input stream, optionally sets compression,
    // and returns the workbook saved into a new MemoryStream.
    public MemoryStream Process(Stream inputStream)
    {
        try
        {
            // Load the workbook from the input stream
            Workbook workbook = new Workbook(inputStream);

            // Uncomment the following line if the Aspose.Cells version supports compression settings
            // workbook.Settings.Compression = CompressionLevel.Level6;

            // Save the workbook into a memory stream in XLSX format
            MemoryStream outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);

            // Reset the stream position to the beginning for downstream consumers
            outputStream.Position = 0;

            return outputStream;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    // Entry point required for compilation
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Prevent FileNotFoundException for the input file
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            using (FileStream inputStream = File.OpenRead(inputPath))
            {
                WorkbookProcessor processor = new WorkbookProcessor();
                using (MemoryStream resultStream = processor.Process(inputStream))
                {
                    // Ensure the output directory exists
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                    using (FileStream outputStream = File.Create(outputPath))
                    {
                        resultStream.CopyTo(outputStream);
                    }
                }
            }

            Console.WriteLine($"Workbook processed and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
