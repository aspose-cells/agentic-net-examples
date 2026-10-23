// Title: Filter and list custom document properties of DateTime type in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# filter custom document properties by DateTime | How to retrieve only DateTime custom properties from an XLSX file in .NET | List Excel custom document properties that contain DateTime values using Aspose.Cells | C# code example for enumerating DateTime custom properties in a workbook | Extract DateTime metadata from custom properties with Aspose.Cells API
// Tags: filter DateTime custom properties Aspose.Cells | enumerate Excel custom document properties C# | retrieve DateTime metadata from workbook | list custom properties by type Aspose.Cells | Aspose.Cells custom property DateTime extraction

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses its CustomDocumentProperties collection, and prints the name and DateTime value of each property whose value is a DateTime, including error handling for missing files.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of custom document properties
            var customProps = workbook.CustomDocumentProperties;

            // Iterate through all custom properties
            foreach (var prop in customProps)
            {
                // Output only properties whose value is a DateTime
                if (prop.Value is DateTime dateValue)
                {
                    Console.WriteLine($"{prop.Name}: {dateValue}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
