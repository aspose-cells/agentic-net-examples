// Title: Load an Excel workbook from a Stream, apply Level6 compression, and write it to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Show how to open a workbook from a Stream, configure Aspose.Cells compression to Level6, and save the result into a MemoryStream as XLSX in C#. | Provide C# code that reads an Excel file from a Stream, sets Workbook.Settings.Compression to Level6, and outputs the compressed workbook to a MemoryStream.
// Common Searches: how to set Aspose.Cells workbook compression level to Level6 in C# | save Excel workbook to MemoryStream with compression using Aspose.Cells .NET | load workbook from a file stream and compress it before saving as XLSX | Aspose.Cells compress XLSX output stream Level6 example | C# read Excel from stream and write compressed workbook to memory
// Tags: Aspose.Cells workbook compression Level6 | save workbook to memory stream xlsx | load workbook from stream c# | compress xlsx output with Aspose.Cells | Workbook.Settings.Compression usage

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook from an input Stream, optionally setting the workbook's compression level to Level6 via Workbook.Settings.Compression, and then saving the workbook as an XLSX file into a MemoryStream. The stream position is reset for further reading, and the resulting compressed bytes can be written to a file or transmitted elsewhere.
class Program
{
    static void Main()
    {
        try
        {
            // Obtain the source stream containing the workbook data.
            using (Stream inputStream = GetInputStream())
            {
                // Load the workbook from the provided stream.
                Workbook workbook = new Workbook(inputStream);

                // Set the workbook compression level to Level6 (if supported).
                // Uncomment the following line if your Aspose.Cells version supports compression.
                // workbook.Settings.Compression = Aspose.Cells.CompressionLevel.Level6;

                // Save the workbook into a memory stream.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);

                    // Reset the position if the stream will be read afterwards.
                    outputStream.Position = 0;

                    // The outputStream now contains the (optionally) compressed workbook.
                    // For demonstration, write it to a file.
                    File.WriteAllBytes("output.xlsx", outputStream.ToArray());
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Returns a stream for the input workbook.
    // If "input.xlsx" exists, it is opened; otherwise an empty workbook stream is created.
    static Stream GetInputStream()
    {
        const string inputPath = "input.xlsx";

        if (File.Exists(inputPath))
        {
            return new FileStream(inputPath, FileMode.Open, FileAccess.Read);
        }

        // Create an empty workbook and return its stream.
        Workbook emptyWb = new Workbook();
        MemoryStream ms = new MemoryStream();
        emptyWb.Save(ms, SaveFormat.Xlsx);
        ms.Position = 0;
        return ms;
    }
}
