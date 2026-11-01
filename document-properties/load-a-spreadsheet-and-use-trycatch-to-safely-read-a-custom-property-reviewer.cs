// Title: Safely read the custom document property "Reviewer" from an Excel workbook using Aspose.Cells for .NET (C#) with try‑catch
// Common Searches: C# Aspose.Cells example to get a custom document property named Reviewer | how to use try‑catch when reading Excel custom properties with Aspose.Cells | retrieve a missing custom property without crashing in Aspose.Cells .NET | read custom document property safely from a workbook using Aspose.Cells | exception handling for custom Excel properties in C# Aspose.Cells
// Tags: Aspose.Cells custom property retrieval C# | try‑catch workbook property access .NET | Excel custom metadata reading Aspose.Cells | handle missing custom property Aspose.Cells | load workbook and access custom fields C#

using System;
using Aspose.Cells;

// The code loads an Excel file into an Aspose.Cells Workbook, accesses its CustomDocumentProperties collection, attempts to read the "Reviewer" property, and uses a try‑catch block to gracefully handle errors such as the property being absent.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        try
        {
            // Access the custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Attempt to read the "Reviewer" property safely
            string reviewer = customProps["Reviewer"]?.ToString();

            Console.WriteLine("Reviewer: " + reviewer);
        }
        catch (Exception ex)
        {
            // Handle any errors (e.g., property not found)
            Console.WriteLine("Error reading Reviewer property: " + ex.Message);
        }
    }
}
