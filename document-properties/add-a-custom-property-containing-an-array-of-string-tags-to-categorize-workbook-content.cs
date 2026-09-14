// Title: Add a comma‑separated 'Tags' custom document property to an Aspose.Cells workbook using C#
// AI Prompts: Create a new Workbook, convert a string array of tags into a single comma‑separated string, and add it as a custom document property named 'Tags' with Aspose.Cells. | Verify that the target output directory exists, create it if missing, then save the workbook to a .xlsx file while handling any exceptions. | Implement error handling for both adding the custom document property and saving the workbook, logging exception details to the console.
// Common Searches: how to add a custom document property called tags in Aspose.Cells C# | store multiple tag values in a single Excel custom property using Aspose.Cells | Aspose.Cells C# add comma separated metadata to workbook | create workbook and set custom property array workaround Aspose.Cells | ensure output folder exists before saving workbook Aspose.Cells C#
// Tags: custom document property comma separated values Aspose.Cells | C# workbook tags metadata implementation | add custom property to Excel file using Aspose.Cells | handle custom property addition errors Aspose.Cells | save workbook with custom metadata Aspose.Cells C# | directory creation before workbook save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, joins a string array of tags into a comma‑separated value, adds it as a custom document property named 'Tags', ensures the output directory exists, and saves the workbook to TaggedWorkbook.xlsx while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define tags to categorize the workbook content
                string[] tags = new string[] { "Finance", "Q1", "Report" };
                // Combine tags into a single string (custom properties do not support arrays directly)
                string tagsValue = string.Join(",", tags);

                // Add a custom document property named "Tags"
                // Use the overload that accepts name and value (type inferred from the value)
                try
                {
                    workbook.CustomDocumentProperties.Add("Tags", tagsValue);
                }
                catch (Exception propEx)
                {
                    Console.WriteLine($"Error adding custom property: {propEx.Message}");
                }

                // Define output path and ensure the directory exists (if any)
                string outputPath = "TaggedWorkbook.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
