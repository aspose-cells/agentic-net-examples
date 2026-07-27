using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReadAuthorProperty
    {
        public static void Run()
        {
            // Path to the Excel file
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Retrieve the Author built‑in property
                string author = workbook.BuiltInDocumentProperties.Author;

                // Output the author name
                Console.WriteLine("Author: " + author);
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ReadAuthorProperty.Run();
        }
    }
}