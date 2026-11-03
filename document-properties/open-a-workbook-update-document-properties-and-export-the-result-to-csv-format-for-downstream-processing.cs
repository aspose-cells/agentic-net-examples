// Title: Update Excel built‑in document properties and save the first worksheet as UTF‑8 CSV using Aspose.Cells for .NET
// Common Searches: how to change author and title properties of an XLSX file with Aspose.Cells in C# | saving a worksheet to UTF‑8 encoded CSV using Aspose.Cells TxtSaveOptions | export first sheet of a workbook to CSV after updating metadata with Aspose.Cells | C# code to set built‑in document properties and convert Excel to CSV | Aspose.Cells example for modifying document properties and exporting to CSV
// Tags: built‑in document properties modification Aspose.Cells | CSV export with UTF‑8 encoding Aspose.Cells | TxtSaveOptions for CSV C# | first worksheet to CSV Aspose.Cells | Excel metadata update .NET

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;
using System.Text;

// The sample loads an existing XLSX workbook, updates its built‑in Author, Title, and Subject properties, configures TxtSaveOptions for UTF‑8 CSV output, and saves the first worksheet as a CSV file, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook from a file
            Workbook workbook = new Workbook(inputPath);

            // Update built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Updated Report";
            workbook.BuiltInDocumentProperties.Subject = "Sales Data";

            // Configure CSV save options (UTF‑8 encoding)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Encoding = Encoding.UTF8
            };

            // Export the first worksheet to CSV format
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook saved as CSV to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
