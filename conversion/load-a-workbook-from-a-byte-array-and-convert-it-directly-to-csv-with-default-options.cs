// Title: Convert an Excel workbook from a byte array to CSV using Aspose.Cells in C#
// Common Searches: Aspose.Cells convert Excel byte array to CSV in C# without writing to disk | How to load an Excel file from a MemoryStream and export as CSV with Aspose.Cells | C# example for reading .xlsx from byte[] and generating .csv using Aspose.Cells | In‑memory Excel to CSV conversion using Aspose.Cells SaveFormat.Csv | Aspose.Cells C# code to transform workbook loaded from byte[] into CSV stream
// Tags: Aspose.Cells load workbook from byte array | Aspose.Cells save workbook as CSV | in-memory Excel to CSV conversion C# | Workbook.Save with SaveFormat.Csv | MemoryStream Excel conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToCsvApp
{
    // Loads an Excel workbook from a byte array via a MemoryStream, saves it directly to CSV using Aspose.Cells default settings into another MemoryStream, and returns the CSV data as a byte array.
    public class ExcelToCsvConverter
    {
        /// <param name="excelBytes">Byte array representing the source Excel file.</param>
        /// <returns>Byte array containing the generated CSV data.</returns>
        public static byte[] ConvertToCsv(byte[] excelBytes)
        {
            try
            {
                // Load the workbook from the provided byte array using a memory stream.
                using (var inputStream = new MemoryStream(excelBytes))
                {
                    var workbook = new Workbook(inputStream);

                    // Save the workbook directly to CSV format into another memory stream.
                    using (var outputStream = new MemoryStream())
                    {
                        workbook.Save(outputStream, SaveFormat.Csv);
                        return outputStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error converting Excel to CSV: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Expect first argument as input Excel file path, second as output CSV path.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExcelToCsvApp <input.xlsx> <output.csv>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Verify input file exists to prevent FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Read Excel file into byte array.
                byte[] excelBytes = File.ReadAllBytes(inputPath);

                // Convert to CSV.
                byte[] csvBytes = ExcelToCsvConverter.ConvertToCsv(excelBytes);

                // Write CSV data to output file.
                File.WriteAllBytes(outputPath, csvBytes);
                Console.WriteLine($"CSV file generated successfully at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Operation failed: {ex.Message}");
            }
        }
    }
}
