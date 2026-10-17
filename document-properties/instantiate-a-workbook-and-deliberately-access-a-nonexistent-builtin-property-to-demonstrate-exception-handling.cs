// Title: C# example: handling exception when accessing a non‑existent built‑in document property in Aspose.Cells Workbook
// Common Searches: Aspose.Cells C# catch exception for missing built‑in document property | how to safely read a built‑in property that may not exist using Aspose.Cells | example of accessing non‑existent workbook built‑in property with try‑catch in .NET | Aspose.Cells built‑in document properties error handling tutorial | retrieve built‑in property with exception handling in Aspose.Cells for .NET
// Tags: exception handling for missing built‑in document property Aspose.Cells | Aspose.Cells workbook built‑in property access error | C# try‑catch Aspose.Cells document properties | Aspose.Cells built‑in document properties missing key handling | Aspose.Cells .NET workbook property retrieval exception

using System;
using Aspose.Cells;

// The sample creates a new Workbook, attempts to read a built‑in document property named "NonExistentProperty", and uses a try‑catch block to capture the resulting exception, demonstrating proper error handling for absent built‑in properties in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a built‑in property that does not exist
            var prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];
            // If no exception occurs, display the (unexpected) value
            Console.WriteLine($"Property value: {prop.Value}");
        }
        catch (Exception ex)
        {
            // Expected exception handling for a missing property
            Console.WriteLine("Exception caught while accessing non‑existent built‑in property:");
            Console.WriteLine(ex.Message);
        }
    }
}
