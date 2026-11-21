// Title: Load an Excel file into an Aspose.Cells Workbook and list its ContentTypeProperties in C#
// Common Searches: how to read custom content type properties from an Excel workbook using Aspose.Cells C# | Aspose.Cells C# enumerate workbook ContentTypeProperties collection | C# load Excel file and display content type metadata with Aspose.Cells | check if Excel file exists before loading with Aspose.Cells .NET | handle exceptions when opening a workbook in Aspose.Cells C#
// Tags: load workbook from Excel file Aspose.Cells | enumerate ContentTypeProperties Aspose.Cells | read custom content type metadata C# | validate Excel file existence before loading | exception handling for workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the specified Excel file exists, loads it into an Aspose.Cells Workbook, accesses the ContentTypeProperties collection, and prints each property's name and value while handling possible runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the collection of content type properties.
            var properties = workbook.ContentTypeProperties;

            // Enumerate and display each property's name and value.
            foreach (var prop in properties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
