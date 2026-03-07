using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToJsonDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source CSV file
            string sourceCsv = "sample.csv";

            // Path for the resulting JSON file
            string outputJson = "sample.json";

            // Create a temporary CSV file for demonstration purposes
            // (In real scenarios the file would already exist)
            File.WriteAllText(sourceCsv, "Name,Age,City\nJohn,30,New York\nAlice,25,London");

            // Convert the CSV file to JSON using Aspose.Cells ConversionUtility.
            // This method uses default load and save options.
            ConversionUtility.Convert(sourceCsv, outputJson);

            // Output the result path
            Console.WriteLine($"CSV file \"{sourceCsv}\" has been converted to JSON file \"{outputJson}\".");

            // Optional: display the generated JSON content
            if (File.Exists(outputJson))
            {
                string jsonContent = File.ReadAllText(outputJson);
                Console.WriteLine("Generated JSON:");
                Console.WriteLine(jsonContent);
            }

            // Clean up temporary files (optional)
            // File.Delete(sourceCsv);
            // File.Delete(outputJson);
        }
    }
}