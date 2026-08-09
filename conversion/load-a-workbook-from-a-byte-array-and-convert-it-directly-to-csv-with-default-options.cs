// Title: C# – Convert an Excel byte[] to CSV with Aspose.Cells (in‑memory)
// Description: Loads an Excel workbook from a byte array using a MemoryStream, then saves it directly to CSV with Aspose.Cells default SaveFormat.Csv settings, returning the CSV as a byte array—no temporary files required.
// Keywords: Aspose.Cells C# | Excel to CSV conversion | byte array Excel | MemoryStream workbook load | SaveFormat.Csv default | .NET in‑memory conversion | no temporary files CSV export | Aspose.Cells API example
// Common Searches: Aspose.Cells convert byte[] to CSV | C# load Excel from MemoryStream and save as CSV | Excel to CSV without saving to disk Aspose | default CSV export Aspose.Cells .NET | in‑memory Excel to CSV conversion C#
// Developer Intent: Export an Excel workbook supplied as a byte array to CSV using Aspose.Cells with default options.
// Use Cases: Web API endpoint that receives an uploaded Excel file (byte[]) and returns CSV data instantly. | Background service that batch‑processes Excel byte streams into CSV for data migration. | Real‑time analytics pipeline that transforms in‑memory Excel reports to CSV for downstream consumption.
// AI Prompts: Generate a C# method that takes a byte[] of an Excel file and returns a byte[] of CSV using Aspose.Cells default settings. | Show how to stream an Excel workbook from a MemoryStream to a CSV response in ASP.NET Core with Aspose.Cells. | Explain best practices for converting large Excel files to CSV in memory with Aspose.Cells to minimize memory footprint.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook from a byte array using a MemoryStream, then saves it directly to CSV with Aspose.Cells default SaveFormat.Csv settings, returning the CSV as a byte array—no temporary files required.
public class CsvConverter
{
    /// <param name="excelBytes">Byte array containing the Excel file.</param>
    /// <returns>Byte array containing the CSV representation.</returns>
    public static byte[] ConvertToCsv(byte[] excelBytes)
    {
        // Load the workbook from the byte array using a MemoryStream (lifecycle rule)
        using (MemoryStream inputStream = new MemoryStream(excelBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Save the workbook directly to CSV using the default options (save rule)
            using (MemoryStream csvStream = new MemoryStream())
            {
                workbook.Save(csvStream, SaveFormat.Csv);
                // Return the CSV data as a byte array
                return csvStream.ToArray();
            }
        }
    }

    // Example usage
    public static void Main()
    {
        // Assume we have an Excel file as a byte array (replace with actual data)
        byte[] excelData = File.ReadAllBytes("sample.xlsx");

        byte[] csvData = ConvertToCsv(excelData);

        // Write the CSV output to a file for verification
        File.WriteAllBytes("output.csv", csvData);

        Console.WriteLine("Conversion to CSV completed successfully.");
    }
}
