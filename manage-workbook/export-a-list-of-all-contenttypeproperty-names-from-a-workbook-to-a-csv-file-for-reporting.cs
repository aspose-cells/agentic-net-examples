// Title: Export workbook ContentTypeProperty names to a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: how to list ContentTypeProperty names from an Excel workbook using Aspose.Cells C# | Aspose.Cells export custom content type properties to CSV in .NET | C# code to write workbook ContentTypeProperties collection to a CSV file | save Excel content type property names as CSV with Aspose.Cells | extract ContentTypeProperty names from .xlsx using Aspose.Cells for .NET
// Tags: Aspose.Cells ContentTypeProperties export to CSV | C# write workbook metadata to CSV | extract custom content type property names .NET | list Excel content type properties using Aspose.Cells | save workbook property names as CSV file

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it with Aspose.Cells, iterates over the workbook's ContentTypeProperties collection, and writes each property's Name to a CSV file (ContentTypeProperties.csv) with a header row, handling quote escaping.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path for the CSV output
            string csvPath = "ContentTypeProperties.csv";

            // Write property names to CSV
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                // CSV header
                writer.WriteLine("PropertyName");

                // Iterate through content type properties
                foreach (var prop in workbook.ContentTypeProperties)
                {
                    // Escape double quotes in the property name
                    string name = prop.Name?.Replace("\"", "\"\"");

                    // Write the name as a quoted CSV field
                    writer.WriteLine($"\"{name}\"");
                }
            }

            Console.WriteLine("Content type property names have been exported to CSV.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
