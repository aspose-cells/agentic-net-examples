// Title: How to export an Excel workbook to a UTF-16 encoded CSV file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save workbook as CSV with UTF-16 encoding | Convert .xlsx to Unicode CSV using TxtSaveOptions in .NET | Export Excel to CSV Unicode (UTF-16) with Aspose.Cells API | C# code to write CSV file in UTF-16 from an Excel workbook | Aspose.Cells CSV export preserving wide characters in C#
// Tags: TxtSaveOptions UTF-16 CSV Aspose.Cells | C# Unicode CSV export from Excel | Aspose.Cells CSV encoding configuration | .NET workbook to UTF-16 CSV conversion | Unicode CSV file generation with Aspose

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an existing .xlsx file with Aspose.Cells, configures TxtSaveOptions to use UTF-16 (Unicode) encoding for CSV output, and saves the workbook as a CSV file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password is supplied; adjust LoadOptions if needed)
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options to use UTF-16 (Unicode) encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.Unicode // UTF-16
            };

            // Export the workbook to CSV with the specified encoding
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime errors (e.g., password protection, I/O issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
