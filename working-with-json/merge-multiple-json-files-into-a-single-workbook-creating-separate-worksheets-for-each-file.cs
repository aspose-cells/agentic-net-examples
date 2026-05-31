using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonMergeDemo
{
    public static class JsonFilesMerger
    {
        /// <summary>
        /// Merges multiple JSON files into a single workbook.
        /// </summary>
        /// <param name="jsonFilePaths">Array of JSON file paths to merge.</param>
        /// <param name="outputPath">Path of the resulting Excel file.</param>
        public static void MergeJsonFiles(string[] jsonFilePaths, string outputPath)
        {
            try
            {
                // Create an empty destination workbook.
                Workbook destWorkbook = new Workbook();

                foreach (string jsonPath in jsonFilePaths)
                {
                    // Ensure the JSON file exists before loading.
                    if (!File.Exists(jsonPath))
                    {
                        Console.WriteLine($"Warning: File not found - {jsonPath}. Skipping.");
                        continue;
                    }

                    // Load the JSON file with MultipleWorksheets enabled.
                    JsonLoadOptions loadOptions = new JsonLoadOptions { MultipleWorksheets = true };
                    Workbook sourceWorkbook = new Workbook(jsonPath, loadOptions);

                    // Merge the source workbook into the destination workbook.
                    destWorkbook.Combine(sourceWorkbook);
                }

                // Save the merged workbook.
                destWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Merged workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
        }

        // Example usage
        public static void RunDemo()
        {
            string[] jsonFiles = new string[]
            {
                "Data1.json",
                "Data2.json",
                "Data3.json"
            };

            string outputFile = "MergedResult.xlsx";

            MergeJsonFiles(jsonFiles, outputFile);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                JsonFilesMerger.RunDemo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}