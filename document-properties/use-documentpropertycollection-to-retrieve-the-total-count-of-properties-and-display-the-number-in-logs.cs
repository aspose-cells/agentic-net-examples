// Title: How to count and log custom document properties in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# count custom document properties in an Excel file | Get total number of custom properties from a workbook with Aspose.Cells | Log custom document property count using Aspose.Cells .NET | Retrieve custom document property collection size from .xlsx in C# | How to read custom document properties count with Aspose.Cells
// Tags: Aspose.Cells count custom document properties | C# retrieve custom document property collection size | log custom document properties count Aspose.Cells | Excel workbook custom properties enumeration .NET | DocumentPropertyCollection Count method usage

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, accesses its CustomDocumentProperties collection via Aspose.Cells, obtains the total count of custom properties, and writes the count to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook file
            string filePath = "sample.xlsx";

            // Ensure the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access custom document properties directly from the workbook
            var customProps = workbook.CustomDocumentProperties;

            // Get the total number of custom properties
            int totalProperties = customProps.Count;

            // Output the count
            Console.WriteLine($"Total custom document properties: {totalProperties}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
