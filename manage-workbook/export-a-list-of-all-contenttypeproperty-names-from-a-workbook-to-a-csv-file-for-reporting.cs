// Title: Export all ContentTypeProperty names from an Excel workbook to a CSV file using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that loads an .xlsx file with Aspose.Cells, iterates over Workbook.ContentTypeProperties, and writes each property Name to a CSV file with a header row. | Show how to verify the input workbook exists, catch exceptions, and export the list of content type property names to a CSV using StreamWriter. | Generate code that escapes commas in property names and saves the output to "ContentTypeNames.csv" while using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# export workbook ContentTypeProperties to CSV | How to list custom content type property names from an Excel file using Aspose.Cells | C# write Excel content type property names into a CSV file | Extract workbook metadata (ContentTypeProperty) with Aspose.Cells .NET | Save Aspose.Cells ContentTypeProperty collection to a CSV report
// Tags: Aspose.Cells export ContentTypeProperty names to CSV | C# iterate Workbook.ContentTypeProperties collection | extract Excel custom content type properties using .NET | StreamWriter write property names to CSV | handle missing workbook file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that loads an Excel workbook with Aspose.Cells, checks the file's existence, iterates through its ContentTypeProperties collection, and writes each property's Name to a CSV file named ContentTypeNames.csv, including a header row and basic error handling.
class ExportContentTypeNames
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path for the CSV output
            string csvPath = "ContentTypeNames.csv";

            // Write content type property names to CSV
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                // Optional header
                writer.WriteLine("ContentTypePropertyName");

                // Iterate through all content type properties
                foreach (var prop in workbook.ContentTypeProperties)
                {
                    // Escape commas if needed (names typically don't contain commas)
                    string name = prop.Name?.Replace(",", "\\,");
                    writer.WriteLine(name);
                }
            }

            Console.WriteLine($"Export completed. Names saved to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
