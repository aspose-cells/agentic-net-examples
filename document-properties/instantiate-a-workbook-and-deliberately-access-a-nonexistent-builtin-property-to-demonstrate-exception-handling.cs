// Title: C# – Handle missing built‑in document property exception in Aspose.Cells
// Description: Creates a Workbook, attempts to read a non‑existent built‑in document property, catches the resulting NullReferenceException, logs the error details, and saves the workbook.
// Keywords: Aspose.Cells | C# | built‑in document property | exception handling | NullReferenceException | nonexistent property | Workbook | document properties | error handling
// Common Searches: Aspose.Cells get built‑in document property by name returns null | C# Aspose.Cells catch exception for missing built‑in property | How to check if a built‑in document property exists in Aspose.Cells | Aspose.Cells document property error handling | Access unknown built‑in property Aspose.Cells
// Developer Intent: Show how to safely access a built‑in document property, detect its absence, and handle the resulting exception without breaking the workflow.
// Use Cases: Validate the presence of a built‑in property before reading its Value to avoid NullReferenceException. | Log detailed error information when an invalid property name is supplied. | Continue normal workbook processing (e.g., saving) after handling a missing property error.
// AI Prompts: Provide a C# example using Aspose.Cells that checks for a built‑in document property’s existence before accessing its Value and gracefully handles missing properties. | Explain why accessing .Value on a null DocumentProperty throws a NullReferenceException and suggest alternative patterns for safe property retrieval in Aspose.Cells. | Generate a unit test that verifies an exception is caught when attempting to read a non‑existent built‑in document property from a Workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a Workbook, attempts to read a non‑existent built‑in document property, catches the resulting NullReferenceException, logs the error details, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook instance (lifecycle create rule)
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a built‑in property that does not exist.
            // The indexer returns null for unknown property names.
            DocumentProperty prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];

            // Accessing the Value of a null property will throw a NullReferenceException.
            Console.WriteLine("Property value: " + prop.Value);
        }
        catch (Exception ex)
        {
            // Demonstrate exception handling for the invalid property access.
            Console.WriteLine("Caught an exception while accessing a non‑existent built‑in property:");
            Console.WriteLine("Exception type: " + ex.GetType().Name);
            Console.WriteLine("Message: " + ex.Message);
        }

        // Save the workbook (lifecycle save rule) to complete the normal workflow.
        workbook.Save("DemoWorkbook.xlsx");
    }
}
