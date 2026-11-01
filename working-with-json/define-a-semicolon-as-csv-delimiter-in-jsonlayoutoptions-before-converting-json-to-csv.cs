// Title: Convert JSON to CSV with a semicolon delimiter using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a JSON file into an Aspose.Cells Workbook and saves it as a CSV file using a semicolon as the field separator. | Show how to configure TxtSaveOptions.Separator to ';' when exporting a workbook to CSV in Aspose.Cells. | Demonstrate checking for the existence of the input JSON file before performing the conversion with a custom CSV delimiter in .NET. | Explain how to use LoadOptions with Auto format detection to read JSON into a workbook before exporting to CSV.
// Common Searches: aspocells set csv delimiter to semicolon in c# | c# convert json file to csv using aspocells with custom separator | how to change txtsaveoptions separator property for csv export | load json into workbook and export as csv with semicolon using aspocells | aspocells csv export custom field separator example
// Tags: json to csv conversion aspocells | txtsaveoptions csv separator property | custom csv delimiter aspocells | c# load json workbook | aspocells csv export semicolon

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks that the input JSON file exists, loads it into an Aspose.Cells Workbook using auto format detection, sets TxtSaveOptions.Separator to a semicolon, and saves the workbook as a CSV file with the custom delimiter.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.json";
            string outputPath = "output.csv";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the JSON file into a workbook (auto format detection)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Configure CSV save options to use semicolon as the separator
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
                {
                    Separator = ';'
                };

                // Save the workbook as CSV using the specified separator
                workbook.Save(outputPath, csvOptions);

                Console.WriteLine($"Conversion completed successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
