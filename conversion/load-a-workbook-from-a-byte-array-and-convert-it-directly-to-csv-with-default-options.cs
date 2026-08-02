// Title: C# – Convert Excel byte array to CSV using Aspose.Cells
// Description: Shows how to load a Workbook from a byte[] via MemoryStream and save it directly to CSV with the default SaveFormat.Csv, avoiding any temporary files.
// Keywords: Aspose.Cells | C# | .NET | byte array to CSV | Excel to CSV in memory | Workbook Load Stream | SaveFormat.Csv | in‑memory conversion | no temporary file
// Common Searches: Aspose.Cells convert byte[] to CSV | load Excel from MemoryStream C# | save workbook as CSV without temp file | C# Excel bytes to CSV example | Aspose.Cells default CSV export
// Developer Intent: Load an Excel workbook from a byte array and export it as a CSV file with default options.
// Use Cases: Web API receives uploaded Excel files as byte arrays and needs instant CSV output for downstream processing. | Generate Excel reports in memory and provide a CSV download without writing the intermediate .xlsx to disk. | Batch conversion of Excel BLOBs stored in a database to CSV for analytics or archival.
// AI Prompts: Create a C# method that accepts a byte[] Excel workbook and returns the CSV content as a string using Aspose.Cells. | Add robust error handling to the conversion routine for null inputs, unsupported formats, and file‑system permission issues. | Show how to process a list of Excel byte arrays in parallel, converting each to a separate CSV file while correctly disposing MemoryStream objects.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvConversion
{
    // Shows how to load a Workbook from a byte[] via MemoryStream and save it directly to CSV with the default SaveFormat.Csv, avoiding any temporary files.
    public class CsvConverter
    {
        /// <param name="excelBytes">Byte array containing the Excel file.</param>
        /// <param name="csvFilePath">Path where the CSV file will be saved.</param>
        public static void ConvertExcelBytesToCsv(byte[] excelBytes, string csvFilePath)
        {
            // Create a memory stream from the input byte array.
            using (MemoryStream excelStream = new MemoryStream(excelBytes))
            {
                // Load the workbook from the stream (uses Workbook(Stream) constructor).
                Workbook workbook = new Workbook(excelStream);

                // Save the workbook directly to CSV format (uses Save(string, SaveFormat) method).
                workbook.Save(csvFilePath, SaveFormat.Csv);
            }
        }

        // Example usage
        public static void Main()
        {
            // Sample Excel data: create a workbook, add data, and get its byte array.
            Workbook sampleWorkbook = new Workbook();
            Worksheet ws = sampleWorkbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["A2"].PutValue("John");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Alice");
            ws.Cells["B3"].PutValue(25);

            // Save the sample workbook to a memory stream to obtain the byte array.
            byte[] excelBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                sampleWorkbook.Save(ms, SaveFormat.Xlsx);
                excelBytes = ms.ToArray();
            }

            // Convert the byte array to CSV.
            string outputCsvPath = "output.csv";
            ConvertExcelBytesToCsv(excelBytes, outputCsvPath);

            Console.WriteLine($"Excel data has been converted to CSV and saved at: {outputCsvPath}");
        }
    }
}
