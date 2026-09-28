// Title: Count and Log Custom Document Properties in an Excel Workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds custom document properties, accesses the CustomDocumentProperties collection, reads its Count, and writes the count to the console with proper try‑catch handling. | Show an Aspose.Cells example that uses DocumentPropertyCollection.Count to determine how many custom properties exist in a workbook and logs the result.
// Common Searches: aspnet count custom document properties in excel using aspose.cells | how to get total number of custom properties from a workbook with Aspose.Cells C# | retrieve and log workbook custom property count Aspose.Cells .NET
// Tags: Aspose.Cells workbook custom property count | C# log property collection size Aspose.Cells | DocumentPropertyCollection size retrieval Aspose | error handling for property count Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new Workbook, adds two custom document properties, accesses the CustomDocumentProperties collection, obtains its Count, and outputs the total count to the console inside a try‑catch block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add custom document properties
            workbook.CustomDocumentProperties.Add("Author", "John Doe");
            workbook.CustomDocumentProperties.Add("Version", 1);

            // Retrieve the collection of custom document properties
            var properties = workbook.CustomDocumentProperties;

            // Get the total count of properties
            int totalCount = properties.Count;

            // Log the count to the console
            Console.WriteLine($"Total document properties count: {totalCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
