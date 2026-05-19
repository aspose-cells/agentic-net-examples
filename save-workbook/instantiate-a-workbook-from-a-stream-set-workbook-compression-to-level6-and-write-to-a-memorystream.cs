using System;
using System.IO;
using Aspose.Cells;

public class WorkbookCompressionExample
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        try
        {
            // Create a sample workbook and save it to a memory stream (represents the input stream)
            using (MemoryStream inputStream = new MemoryStream())
            {
                Workbook sampleWorkbook = new Workbook();
                sampleWorkbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                sampleWorkbook.Save(inputStream, SaveFormat.Xlsx);
                inputStream.Position = 0; // Reset for reading

                // Load the workbook from the input stream
                Workbook workbook = new Workbook(inputStream);

                // Configure OOXML save options with Level6 compression
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    CompressionType = OoxmlCompressionType.Level6
                };

                // Save the workbook to an output memory stream using the configured options
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, saveOptions);
                    outputStream.Position = 0; // Ready for further processing

                    // Example: write the compressed workbook to a file (optional)
                    string outputPath = "CompressedWorkbook.xlsx";
                    File.WriteAllBytes(outputPath, outputStream.ToArray());
                    Console.WriteLine($"Compressed workbook saved to '{outputPath}'.");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.Error.WriteLine($"Runtime exception: {ex.Message}");
            throw;
        }
    }
}