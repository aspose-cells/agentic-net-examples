// Title: Demonstrate exception handling when accessing a missing built‑in document property in an Aspose.Cells Workbook (C#)
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, attempts to read a built‑in document property that does not exist, and catches the thrown exception. | Show how to check for the existence of a built‑in property before accessing it and handle invalid property names gracefully using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells how to catch exception for unknown built‑in document property | Aspose.Cells example retrieving built‑in property with try‑catch | What error is thrown when accessing a non‑existent built‑in property in an Aspose.Cells workbook | Sample code for handling missing built‑in document property in Aspose.Cells
// Tags: Aspose.Cells built‑in document property exception | C# workbook missing property handling | Aspose.Cells safe built‑in property access | exception handling Aspose.Cells document properties | Aspose.Cells retrieve unknown built‑in property

using System;
using Aspose.Cells;

// // Example showing how to instantiate a Workbook, attempt to read a non‑existent built‑in document property, and catch the resulting exception.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a non‑existent built‑in property
            var prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];
            Console.WriteLine($"Property value: {prop.Value}");
        }
        catch (Exception ex)
        {
            // Expected exception handling
            Console.WriteLine("Exception caught while accessing non‑existent property:");
            Console.WriteLine(ex.Message);
        }
    }
}
