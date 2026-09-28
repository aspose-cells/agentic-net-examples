// Title: Convert a JSON file to CSV with a semicolon delimiter using Aspose.Cells in C#
// AI Prompts: Generate C# code that reads a JSON file, loads it into an Aspose.Cells Workbook with auto‑detect, and saves the first worksheet as a CSV using a semicolon as the separator. | Show how to configure TxtSaveOptions in Aspose.Cells to export workbook data to CSV with a custom delimiter in a .NET console application.
// Common Searches: aspocells c# convert json file to csv with custom separator | how to set semicolon delimiter when saving workbook as csv using Aspose.Cells | load json data into workbook and export to csv using TxtSaveOptions in .NET | c# Aspose.Cells JsonUtility load json and save as csv with custom delimiter
// Tags: Aspose.Cells JsonUtility load JSON | TxtSaveOptions CSV separator | Workbook.Save CSV with semicolon | C# Aspose.Cells JSON to CSV conversion | LoadOptions auto format detection

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example reads a JSON file, loads its tabular data into an Aspose.Cells Workbook using auto‑detect load options, and then saves the first worksheet as a CSV file where the fields are separated by a semicolon.
class JsonToCsvConverter
{
    static void Main()
    {
        try
        {
            // Path to the source JSON file
            string jsonFilePath = "input.json";

            // Verify that the JSON file exists
            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"Error: JSON file '{jsonFilePath}' not found.");
                return;
            }

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Load JSON data into a workbook using LoadOptions with Auto detection
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent)))
            {
                workbook = new Workbook(ms, loadOptions);
            }

            // Configure CSV save options with a custom delimiter (e.g., semicolon)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ';'   // Custom delimiter
            };

            // Save the first worksheet as a CSV file using the specified options
            string csvOutputPath = "output.csv";
            workbook.Save(csvOutputPath, csvOptions);

            Console.WriteLine($"JSON data has been converted to CSV and saved to '{csvOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
