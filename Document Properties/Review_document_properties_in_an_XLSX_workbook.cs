using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace DocumentPropertiesReview
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX workbook
            string inputPath = "SampleWorkbook.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Review Built‑in Document Properties
            // -------------------------------------------------
            Console.WriteLine("=== Built‑in Document Properties ===");
            // Iterate through all built‑in properties
            foreach (DocumentProperty prop in workbook.BuiltInDocumentProperties)
            {
                // Display property name and its value
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }

            // -------------------------------------------------
            // Review Custom Document Properties
            // -------------------------------------------------
            Console.WriteLine("\n=== Custom Document Properties ===");
            // Iterate through all custom properties
            foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} (Type: {prop.Type})");
            }

            // No modifications are made, but if you wanted to save the workbook
            // after reviewing or making changes, you could use the save rule:
            // workbook.Save("ReviewedWorkbook.xlsx");
        }
    }
}