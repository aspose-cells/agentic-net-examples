// Title: Export a large Excel workbook to CSV with UTF‑32 encoding using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net export excel workbook to csv with utf-32 encoding using aspose.cells | how to configure TxtSaveOptions for UTF-32 CSV output in Aspose.Cells .NET | saving a 10,000‑row worksheet as CSV with custom encoding in C# | Aspose.Cells example for large dataset CSV export with UTF-32 | set UTF-32 encoding for CSV file when using Aspose.Cells SaveFormat.CSV
// Tags: TxtSaveOptions UTF-32 CSV export | Aspose.Cells large dataset CSV conversion | C# SaveFormat.CSV custom encoding | CSV output encoding configuration Aspose.Cells | Export workbook to CSV with Unicode encoding .NET

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample creates a workbook, fills the first worksheet with 10,000 rows of data, sets TxtSaveOptions to CSV format with UTF‑32 encoding, and saves the file as a CSV using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large dataset
            for (int row = 0; row < 10000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                sheet.Cells[row, 1].PutValue(row * 10);
            }

            // Configure CSV save options using TxtSaveOptions (compatible with CSV format)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            csvOptions.Encoding = Encoding.UTF32; // Set UTF‑32 encoding

            // Save the workbook as a CSV file with the specified options
            string outputPath = "LargeDatasetOutput.csv";
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
