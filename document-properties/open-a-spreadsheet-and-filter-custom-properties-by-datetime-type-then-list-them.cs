// Title: Read and display DateTime custom document properties from an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, iterates through its CustomDocumentProperties collection, and prints only the properties whose Type is DateTime, including handling for missing files and runtime errors. | Generate a .NET example that filters custom document properties by DateTime type, retrieves each property's name and value, and writes the results to the console using Aspose.Cells.
// Common Searches: Aspose.Cells C# filter custom document properties by DateTime | How to get DateTime custom properties from an Excel workbook using Aspose.Cells | C# code to list DateTime type custom properties in an .xlsx file with Aspose.Cells | Retrieve custom document properties of type DateTime in .NET Excel processing
// Tags: filter DateTime custom document properties Aspose.Cells | enumerate custom properties Excel .NET | read custom document properties from .xlsx using Aspose.Cells | list DateTime type properties C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (input.xlsx) with Aspose.Cells, accesses the CustomDocumentProperties collection, uses reflection to identify properties whose Type equals "DateTime", and prints each matching property's name and value. It includes checks for file existence and robust error handling for individual property processing.
class Program
{
    static void Main()
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
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Iterate through all custom properties
            foreach (var prop in customProps)
            {
                try
                {
                    // Use reflection to avoid direct dependency on specific types
                    var typeProp = prop.GetType().GetProperty("Type");
                    var nameProp = prop.GetType().GetProperty("Name");
                    var valueProp = prop.GetType().GetProperty("Value");

                    if (typeProp == null || nameProp == null || valueProp == null)
                        continue;

                    var typeValue = typeProp.GetValue(prop)?.ToString();
                    // Check if the property type is DateTime
                    if (typeValue == "DateTime")
                    {
                        var name = nameProp.GetValue(prop);
                        var value = valueProp.GetValue(prop);
                        Console.WriteLine($"Name: {name}, Value: {value}");
                    }
                }
                catch (Exception innerEx)
                {
                    // Handle any errors while processing an individual property
                    Console.WriteLine($"Error processing a property: {innerEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
