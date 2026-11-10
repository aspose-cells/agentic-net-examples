// Title: Export all ContentTypeProperty names from an Excel workbook to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook ContentTypeProperty names to CSV | How to list ContentTypeProperties of an Excel file with Aspose.Cells | Save Excel metadata (ContentTypeProperty) as CSV using Aspose.Cells .NET | Aspose.Cells iterate ContentTypeProperties and write to CSV in C# | Export workbook custom content type property names to a CSV file with Aspose.Cells
// Tags: aspose.cells contenttypeproperty name extraction | enumerate workbook contenttypeproperties with aspose.cells | c# write excel metadata to csv using aspose.cells | aspose.cells workbook property list to csv | export excel custom property names via aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through its ContentTypeProperties collection, and writes each property's Name to a CSV file with a header. It creates the output directory if needed, validates the input file, and handles exceptions.
class ContentTypePropertyExporter
{
    static void Main()
    {
        // Paths for input workbook and output CSV
        string workbookPath = @"C:\Input\SampleWorkbook.xlsx";
        string csvPath = @"C:\Output\ContentTypeProperties.csv";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(csvPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Open a StreamWriter for the CSV file
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                // Write CSV header
                writer.WriteLine("ContentTypePropertyName");

                // Iterate through all ContentTypeProperty objects in the workbook
                foreach (var prop in workbook.ContentTypeProperties)
                {
                    // Write the property name to the CSV (escape double quotes if needed)
                    string name = prop.Name?.Replace("\"", "\"\"");
                    writer.WriteLine($"\"{name}\"");
                }
            }

            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
