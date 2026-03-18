using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Paths for the source CSV file and the target JSON file
            string csvPath = "sample.csv";
            string jsonPath = "sample.json";

            // Create a sample CSV file (this step simulates the existing XLTX CSV document)
            File.WriteAllText(csvPath,
                "Name,Age,City\n" +
                "John,30,New York\n" +
                "Alice,25,London\n" +
                "Bob,35,Paris");

            // Convert the CSV file to JSON using the default conversion settings.
            // The Convert method automatically determines load and save options.
            ConversionUtility.Convert(csvPath, jsonPath);

            Console.WriteLine($"Conversion completed: '{csvPath}' -> '{jsonPath}'");

            // Optional: display the generated JSON content
            if (File.Exists(jsonPath))
            {
                string jsonContent = File.ReadAllText(jsonPath);
                Console.WriteLine("Generated JSON:");
                Console.WriteLine(jsonContent);
            }

            // Clean up temporary files (optional)
            // File.Delete(csvPath);
            // File.Delete(jsonPath);
        }
    }
}