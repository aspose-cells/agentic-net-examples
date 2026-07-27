using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsContentTypeReport
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string csvPath = "ContentTypePropertiesReport.csv";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the directory for the CSV file exists
                string csvDir = Path.GetDirectoryName(csvPath);
                if (!string.IsNullOrEmpty(csvDir) && !Directory.Exists(csvDir))
                {
                    Directory.CreateDirectory(csvDir);
                }

                // Prepare the CSV file for writing the report
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    // Write header
                    writer.WriteLine("PropertyName");

                    // Iterate through custom document properties (used as content‑type properties)
                    var enumerator = workbook.CustomDocumentProperties.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        // Use dynamic to access the Name property without explicit type reference
                        dynamic prop = enumerator.Current;
                        writer.WriteLine(prop.Name);
                    }
                }

                Console.WriteLine($"Content type property names have been exported to '{csvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}