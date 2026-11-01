// Title: Convert an Excel .xlsx workbook to CSV with Aspose.Cells for .NET while preserving cell data types
// Common Searches: C# Aspose.Cells convert .xlsx to .csv keeping numeric formatting | how to export Excel workbook as CSV using TxtSaveOptions with UTF-8 in .NET | Aspose.Cells save workbook to CSV without losing cell type information | custom comma separator CSV export from Excel using Aspose.Cells C# example
// Tags: Aspose.Cells TxtSaveOptions CSV export | C# workbook.Save CSV with UTF-8 encoding | Excel to CSV conversion preserving data types | custom CSV separator configuration Aspose.Cells | SaveFormat.Csv usage in Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program checks for the input .xlsx file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions for CSV output with UTF-8 encoding and a comma separator, and saves the workbook as a CSV file while retaining the original cell data types.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure TXT save options for CSV output
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8, // Use UTF-8 encoding
                Separator = ','           // Standard comma separator
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
