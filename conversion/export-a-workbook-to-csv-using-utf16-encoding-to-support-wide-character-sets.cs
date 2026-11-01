// Title: Export an Excel workbook to a UTF-16 encoded CSV file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to CSV with UTF-16 encoding | Save Excel as Unicode CSV using Aspose.Cells .NET API | Convert XLSX to UTF-16 CSV in C# with Aspose.Cells | Set CSV encoding to UTF-16 when saving with Aspose.Cells SaveOptions | How to generate a Unicode CSV file from an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells CSV UTF-16 export | TxtSaveOptions Unicode CSV | C# Excel to CSV conversion Aspose | Workbook.Save CSV encoding option | Unicode CSV output Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example verifies the source XLSX file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions for CSV format with UTF-16 (Encoding.Unicode) encoding, and saves the workbook as a CSV file while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Set up TXT save options for CSV output with UTF‑16 (Unicode) encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
            csvOptions.Encoding = Encoding.Unicode; // UTF‑16

            // Export the workbook to CSV using the configured options
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook successfully saved as CSV to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
