// Title: Retrieve and log the total number of custom document properties using Aspose.Cells for .NET (C#)
// Common Searches: C# code to get the number of custom document properties in an Excel workbook with Aspose.Cells | how to count custom document properties in a .NET Excel file using Aspose.Cells | log the count of workbook custom properties in a C# console application | Aspose.Cells API to retrieve custom document properties count | example of accessing CustomDocumentProperties.Count in C#
// Tags: Aspose.Cells custom document properties count | C# retrieve workbook custom properties | log custom document properties Aspose.Cells | Aspose.Cells CustomDocumentProperties enumeration | Excel custom properties count .NET

using Aspose.Cells;
using System;

// The example creates or loads a Workbook, accesses its CustomDocumentProperties collection, reads the Count property to determine how many custom properties exist, and writes that number to the console using C# and Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the collection of custom document properties
            var properties = workbook.CustomDocumentProperties;

            // Retrieve the total count of properties
            int totalCount = properties.Count;

            // Log the count
            Console.WriteLine($"Total document properties count: {totalCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
