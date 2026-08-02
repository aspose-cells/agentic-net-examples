using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCloneBuiltInProperties
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook (source)
                string templatePath = "TemplateWorkbook.xlsx";

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook sourceWorkbook = new Workbook(templatePath);

                // Create a new empty workbook (destination)
                Workbook destinationWorkbook = new Workbook();

                // Clone built‑in document properties from source to destination
                foreach (var sourceProp in sourceWorkbook.BuiltInDocumentProperties)
                {
                    // Destination already contains the same built‑in properties; set the value directly.
                    destinationWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
                }

                // Save the destination workbook with the cloned properties
                string outputPath = "ClonedPropertiesWorkbook.xlsx";
                destinationWorkbook.Save(outputPath);

                Console.WriteLine($"Built‑in properties cloned and workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}