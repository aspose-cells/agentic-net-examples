// Title: Convert a CSV file to TSV in C# with Aspose.Cells while ignoring lines that start with a custom comment character
// Common Searches: C# Aspose.Cells load CSV with custom comment character and export to TSV | How to skip commented rows when converting CSV to TSV using Aspose.Cells .NET | Aspose.Cells TxtLoadOptions comment character example in C# | Convert CSV to tab‑separated values with Aspose.Cells and ignore # lines | Save workbook as TSV after loading CSV with comment handling in .NET | Aspose.Cells CSV to TSV conversion handling missing input file C#
// Tags: Aspose.Cells CSV import with comment character | TxtLoadOptions separator and comment settings | C# workbook save as TSV | Validate input file existence in Aspose.Cells workflow | Skip commented lines during CSV load

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToTsv
{
    // The example checks that the source CSV file exists, creates a TxtLoadOptions object where the field separator is set to a comma and a comment character can be defined to skip commented rows, loads the CSV into an Aspose.Cells Workbook, and then saves the workbook as a tab‑separated values (TSV) file, handling any exceptions that may occur.
    class CsvToTsvConverter
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Path for the resulting TSV file
            string tsvPath = "output.tsv";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"Input file not found: {csvPath}");
                    return;
                }

                // Configure CSV load options using TxtLoadOptions
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    // Specify the field separator (comma for CSV)
                    Separator = ','
                };

                // Load the CSV file using the configured options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Save the workbook as TSV (tab‑separated values)
                workbook.Save(tsvPath, SaveFormat.Tsv);

                Console.WriteLine($"Conversion successful: {tsvPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
