// Title: C# – Convert Aspose.Cells Workbook to CSV and GZIP‑compress the output
// Description: Creates a workbook, fills the first worksheet, saves it to a MemoryStream in CSV format (SaveFormat.Csv), then streams the CSV data into a GZipStream to produce a compressed file (output.csv.gz). All streams are properly disposed.
// Keywords: Aspose.Cells CSV export | C# GZipStream | gzip compressed CSV | Aspose.Cells SaveFormat.Csv | compress Excel to .gz | .NET CSV compression | memory stream to gzip | export workbook as gzipped CSV
// Common Searches: Aspose.Cells save workbook as CSV C# | gzip CSV file in C# using Aspose.Cells | compress Excel CSV output with GZipStream .NET | create gzipped CSV from workbook programmatically | export multiple worksheets to gzipped CSV C#
// Developer Intent: Export a workbook to CSV and write the result directly into a .gz archive.
// Use Cases: Automated reporting pipelines that need lightweight CSV files | Data exchange with services that require compressed CSV payloads | Archiving Excel reports as gzipped CSV to save storage space | Streaming CSV data to a compressed file without intermediate disk writes
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook as CSV and compresses it with GZipStream, ensuring all streams are disposed correctly. | Explain how to set a custom CSV delimiter and choose a GZip compression level in the example. | Show how to iterate over all worksheets, creating separate gzipped CSV files for each. | Generate code that writes the CSV using UTF‑8 encoding and applies maximum GZip compression. | Describe a streaming approach for very large workbooks to avoid high memory usage while gzipping the CSV output.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsGzipExample
{
    // Creates a workbook, fills the first worksheet, saves it to a MemoryStream in CSV format (SaveFormat.Csv), then streams the CSV data into a GZipStream to produce a compressed file (output.csv.gz). All streams are properly disposed.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Save the workbook to a memory stream in CSV format (uses Save(Stream, SaveFormat) rule)
            using (MemoryStream csvStream = new MemoryStream())
            {
                workbook.Save(csvStream, SaveFormat.Csv);
                csvStream.Position = 0; // Reset stream position for reading

                // Define the path for the compressed GZIP file
                string gzipFilePath = "output.csv.gz";

                // Compress the CSV data using GZipStream and write to the file system
                using (FileStream fileStream = new FileStream(gzipFilePath, FileMode.Create, FileAccess.Write))
                using (GZipStream gzipStream = new GZipStream(fileStream, CompressionMode.Compress))
                {
                    csvStream.CopyTo(gzipStream);
                }

                Console.WriteLine($"Workbook successfully converted to CSV and compressed to '{gzipFilePath}'.");
            }

            // Clean up the workbook instance
            workbook.Dispose();
        }
    }
}
