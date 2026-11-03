// Title: Convert an Excel workbook from a byte array to CSV with Aspose.Cells for .NET
// Common Searches: C# how to load Excel file from byte[] and export to CSV using Aspose.Cells | Aspose.Cells convert in-memory Excel workbook to CSV without saving to disk | example of saving first worksheet as CSV from MemoryStream with Aspose.Cells | byte array to CSV conversion using Aspose.Cells SaveFormat.Csv in .NET | load workbook from byte array and get CSV bytes with Aspose.Cells API
// Tags: Aspose.Cells load workbook from byte array | Aspose.Cells export worksheet to CSV | C# memory stream Excel to CSV conversion | SaveFormat.Csv usage Aspose.Cells | default CSV options Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookConversionApp
{
    // The ConvertExcelToCsv method reads an Excel file supplied as a byte array via a MemoryStream, creates an Aspose.Cells Workbook, saves the first worksheet directly to CSV using the default SaveFormat.Csv settings into another MemoryStream, and returns the CSV data as a byte array.
    public class WorkbookConverter
    {
        /// <param name="excelBytes">The byte array containing the Excel file.</param>
        /// <returns>A byte array containing the CSV representation of the first worksheet.</returns>
        public static byte[] ConvertExcelToCsv(byte[] excelBytes)
        {
            try
            {
                // Load the workbook from the provided byte array.
                using (var inputStream = new MemoryStream(excelBytes))
                {
                    var workbook = new Workbook(inputStream);

                    // Save the first worksheet as CSV into a memory stream.
                    using (var outputStream = new MemoryStream())
                    {
                        workbook.Save(outputStream, SaveFormat.Csv);
                        return outputStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: WorkbookConverter <inputExcelPath> <outputCsvPath>");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];

                // Verify input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Read Excel file bytes.
                byte[] excelBytes = File.ReadAllBytes(inputPath);

                // Convert to CSV.
                byte[] csvBytes = WorkbookConverter.ConvertExcelToCsv(excelBytes);

                // Write CSV bytes to output file.
                File.WriteAllBytes(outputPath, csvBytes);
                Console.WriteLine($"CSV file saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
