// Title: How to convert an Excel workbook to a UTF-8 CSV file and compress it into a .gz archive using Aspose.Cells for .NET
// Common Searches: C# convert .xlsx to CSV with UTF-8 encoding using Aspose.Cells | Create a .csv.gz file from an Excel workbook in .NET | Compress CSV output from Aspose.Cells with GZipStream | Save workbook as CSV then zip it programmatically in C# | Handle missing Excel file and password protection when exporting to CSV with Aspose.Cells
// Tags: Aspose.Cells CSV export UTF-8 | gzip compression of CSV files in .NET | C# workbook to gzip archive conversion | remove temporary CSV after compression | password protected workbook handling Aspose.Cells

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an XLSX workbook with Aspose.Cells, saves it as a UTF-8 CSV file, compresses the CSV into a .gz archive using GZipStream, deletes the temporary CSV, and includes error handling for missing files and password‑protected workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input, intermediate CSV, and final GZIP file
            string inputPath = "input.xlsx";
            string csvPath = "output.csv";
            string gzipPath = "output.csv.gz";

            try
            {
                // Ensure the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no password handling)
                Workbook workbook = new Workbook(inputPath);

                // Save workbook as CSV with UTF‑8 encoding
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
                {
                    Encoding = System.Text.Encoding.UTF8
                };
                workbook.Save(csvPath, csvOptions);

                // Compress the CSV file into GZIP
                using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
                using (FileStream gzipStream = new FileStream(gzipPath, FileMode.Create, FileAccess.Write))
                using (GZipStream compressionStream = new GZipStream(gzipStream, CompressionMode.Compress))
                {
                    csvStream.CopyTo(compressionStream);
                }

                // Delete the temporary CSV file
                if (File.Exists(csvPath))
                {
                    File.Delete(csvPath);
                }

                Console.WriteLine($"GZIP file created successfully: {gzipPath}");
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                Console.WriteLine("The workbook is password protected. Provide a password or use an unprotected file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
