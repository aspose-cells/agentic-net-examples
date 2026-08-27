// Title: Convert an Excel workbook from a byte array to CSV using Aspose.Cells for .NET
// AI Prompts: Generate C# code that accepts a byte[] containing an Excel file, loads it into an Aspose.Cells Workbook via MemoryStream, and saves it as a CSV file using the default SaveFormat.Csv options. | Write a method that reads Excel bytes, creates a Workbook object, and exports the data directly to a CSV file without creating intermediate files, leveraging Aspose.Cells' Save method.
// Common Searches: how to convert excel byte array to csv with aspose.cells in c# | asp.net core load excel from memory stream and save as csv using aspose | c# convert in‑memory workbook to csv default settings aspose.cells | save workbook from byte[] to csv file Aspose.Cells example
// Tags: byte-array to workbook Aspose.Cells | Workbook.Save CSV Aspose.Cells | in-memory Excel CSV export C# | Aspose.Cells memory stream conversion

using System;
using System.IO;
using Aspose.Cells;

// The example reads an Excel file into a byte array, wraps it in a MemoryStream, loads it into an Aspose.Cells Workbook, and saves the workbook directly to a CSV file using the default CSV save options.
public class CsvConversionDemo
{
    // Loads an Excel workbook from a byte array and saves it as CSV using default options.
    public static void ConvertExcelBytesToCsv(byte[] excelBytes, string csvFilePath)
    {
        // Wrap the byte array in a MemoryStream (lifecycle rule: Workbook(Stream) constructor)
        using (MemoryStream stream = new MemoryStream(excelBytes))
        {
            // Load the workbook from the stream
            Workbook workbook = new Workbook(stream);

            // Save the workbook directly to CSV (default options, using Save method with format)
            workbook.Save(csvFilePath, SaveFormat.Csv);
        }
    }

    // Example usage
    public static void Main()
    {
        // Sample Excel data (replace with actual byte array in real scenario)
        byte[] sampleExcel = File.ReadAllBytes("sample.xlsx");

        // Destination CSV file path
        string outputCsv = "output.csv";

        // Perform conversion
        ConvertExcelBytesToCsv(sampleExcel, outputCsv);

        Console.WriteLine($"Workbook converted to CSV at: {outputCsv}");
    }
}
