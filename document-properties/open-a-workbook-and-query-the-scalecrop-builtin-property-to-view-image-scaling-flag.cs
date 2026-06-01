using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class BuiltInDocumentPropertyScaleCropDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string filePath = "input.xlsx";

                // Prevent FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access built‑in document properties
                var properties = workbook.BuiltInDocumentProperties;

                // Get the ScaleCrop property (true = thumbnail is scaled, false = original size)
                bool scaleCrop = properties.ScaleCrop;

                // Display the value
                Console.WriteLine("ScaleCrop property value: " + scaleCrop);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}