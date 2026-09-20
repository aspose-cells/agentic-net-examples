// Title: Convert an Excel workbook to a gzipped CSV file using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, saves it as CSV to a MemoryStream using TxtSaveOptions, and writes the stream to a .gz file with GZipStream. | Show how to configure delimiter and encoding for CSV export in Aspose.Cells before compressing the output with GZipStream. | Provide a .NET console example that converts a workbook to CSV, streams the result, and creates a gzip‑compressed CSV file.
// Common Searches: aspnet convert excel to csv and gzip using aspose.cells | c# save workbook as csv to memory stream then compress with gzipstream | how to use TxtSaveOptions for csv export before gzip compression in .net | generate gzipped csv from xlsx with aspose.cells library | compress csv output from aspose.cells workbook with gzip in C# console app
// Tags: Aspose.Cells CSV export | GZipStream compression of CSV | TxtSaveOptions CSV configuration | MemoryStream to GZipStream .NET | gzipped CSV generation from Excel

using Aspose.Cells;
using System;
using System.IO;
using System.IO.Compression;

// // Loads an Excel workbook via Aspose.Cells, saves it as CSV into a MemoryStream using TxtSaveOptions, then compresses the CSV data into a .gz file with GZipStream.
class Program
{
    static void Main(string[] args)
    {
        // Path to the source Excel workbook
        string inputPath = "input.xlsx";

        // Path for the compressed CSV output
        string outputPath = "output.csv.gz";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook as CSV into a memory stream
        using (MemoryStream csvStream = new MemoryStream())
        {
            // Configure CSV save options if needed (e.g., delimiter, encoding)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            workbook.Save(csvStream, csvOptions);

            // Reset stream position to the beginning before reading
            csvStream.Position = 0;

            // Create the output file and compress the CSV data using GZIP
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            using (GZipStream gzipStream = new GZipStream(fileStream, CompressionMode.Compress))
            {
                csvStream.CopyTo(gzipStream);
            }
        }

        Console.WriteLine("Workbook successfully converted to CSV and compressed as GZIP.");
    }
}
