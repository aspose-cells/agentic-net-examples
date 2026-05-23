using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsCsvGzipDemo
{
    public class Program
    {
        public static void Main()
        {
            // Initialize a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Save the workbook as CSV into a memory stream using the provided Save method
            using (MemoryStream csvStream = new MemoryStream())
            {
                workbook.Save(csvStream, SaveFormat.Csv);
                csvStream.Position = 0; // Reset for reading

                // Create the GZIP output file
                using (FileStream gzipFile = new FileStream("output.csv.gz", FileMode.Create, FileAccess.Write))
                using (GZipStream gzipStream = new GZipStream(gzipFile, CompressionMode.Compress))
                {
                    // Copy CSV data into the GZIP stream
                    csvStream.CopyTo(gzipStream);
                }
            }

            // Clean up workbook resources
            workbook.Dispose();

            Console.WriteLine("Workbook converted to CSV and compressed as GZIP successfully.");
        }
    }
}