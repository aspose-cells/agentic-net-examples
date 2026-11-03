// Title: Convert a TSV file to a comma‑separated CSV using Aspose.Cells LoadOptions and TxtSaveOptions in C#
// Common Searches: C# code to read a .tsv file with Aspose.Cells and save as .csv | Aspose.Cells LoadOptions for TSV format example in .NET | How to set comma delimiter when exporting workbook to CSV with Aspose.Cells | Saving TSV workbook as UTF‑8 CSV using Aspose.Cells TxtSaveOptions | Convert tab‑delimited spreadsheet to comma‑delimited CSV programmatically
// Tags: Aspose.Cells LoadOptions TSV format C# | TxtSaveOptions comma separator CSV | TSV to CSV conversion Aspose.Cells | UTF-8 CSV export Aspose.Cells | Workbook.Save CSV custom delimiter

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The sample checks for an input.tsv file, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Tsv, configures TxtSaveOptions to use a comma separator and UTF‑8 encoding, and saves the content as output.csv while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.tsv";
            const string outputPath = "output.csv";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the TSV file using the TSV load format
            var loadOptions = new LoadOptions(LoadFormat.Tsv);
            var workbook = new Workbook(inputPath, loadOptions);

            // Configure text save options to use commas as separators and UTF-8 encoding
            var txtOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',               // Use comma delimiter
                Encoding = Encoding.UTF8       // Set desired encoding
            };

            // Export the workbook content to a CSV file
            workbook.Save(outputPath, txtOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
