using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsBuiltInPropertiesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be opened
            string filePath = "sample.xlsx";

            try
            {
                // Verify that the file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Retrieve the collection of built‑in document properties
                BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

                // Iterate over each property and display its name and value
                foreach (DocumentProperty prop in builtInProps)
                {
                    Console.WriteLine($"{prop.Name}: {prop.Value}");
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}