using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class QueryScaleCropProperty
    {
        public static void Run()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify that the file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access built‑in document properties
                BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

                // Query the ScaleCrop property (display mode of the document thumbnail)
                bool scaleCrop = properties.ScaleCrop;

                // Output the value
                Console.WriteLine($"ScaleCrop property value: {scaleCrop}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            QueryScaleCropProperty.Run();
        }
    }
}