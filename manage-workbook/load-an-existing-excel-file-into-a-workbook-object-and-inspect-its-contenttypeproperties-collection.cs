// Title: Load an existing .xlsx workbook with Aspose.Cells in C# and list its ContentTypeProperties
// AI Prompts: Write C# code that opens a specified .xlsx file using Aspose.Cells, retrieves the Workbook.ContentTypeProperties collection, and prints each property's name and value. | Show how to safely load an Excel workbook with Aspose.Cells, verify the file exists, and handle exceptions while accessing content‑type metadata. | Demonstrate iterating over ContentTypeProperties via reflection in Aspose.Cells to avoid compile‑time type dependencies.
// Common Searches: how to read custom content type properties from an existing Excel file using Aspose.Cells C# | C# Aspose.Cells enumerate workbook ContentTypeProperties collection | list metadata stored in ContentTypeProperties of a loaded workbook Aspose.Cells | Aspose.Cells reflection example to get property name and value from ContentTypeProperties
// Tags: load workbook contenttypeproperties aspose.cells c# | enumerate contenttypeproperties collection | excel custom metadata extraction aspose.cells | reflection based property access aspose.cells | handle missing excel file aspose.cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of an input .xlsx file, loads it into an Aspose.Cells Workbook, accesses the ContentTypeProperties collection, and uses reflection to output each property's name and value while handling potential runtime errors.
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
            // Load the existing Excel file into a Workbook object
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the collection of content type properties
            var contentProps = workbook.ContentTypeProperties;

            // Iterate through each property using reflection to avoid direct type dependencies
            foreach (var prop in contentProps)
            {
                var propType = prop.GetType();
                var nameProp = propType.GetProperty("Name");
                var valueProp = propType.GetProperty("Value");

                var name = nameProp?.GetValue(prop, null);
                var value = valueProp?.GetValue(prop, null);

                Console.WriteLine($"{name}: {value}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
