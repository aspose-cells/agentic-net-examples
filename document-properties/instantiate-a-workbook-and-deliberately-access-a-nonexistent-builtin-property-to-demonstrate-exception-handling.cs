// Title: Aspose.Cells C# – Exception handling for a missing built‑in document property
// Description: This C# example creates a Workbook, tries to read a non‑existent built‑in document property, deliberately triggers a NullReferenceException, catches the error, displays its type and message, and finally saves the workbook as DemoWorkbook.xlsx.
// Keywords: Aspose.Cells | C# | built‑in document property | missing property | exception handling | NullReferenceException | Workbook.Save | DocumentProperty | error handling | Aspose.Cells for .NET
// Common Searches: how to catch exception when accessing unknown built‑in document property Aspose.Cells | Aspose.Cells C# example for handling missing document property | null reference error reading non‑existent property in Aspose.Cells | save workbook after property access failure Aspose.Cells | validate built‑in document property existence Aspose.Cells .NET
// Developer Intent: Demonstrate catching and reporting an exception caused by accessing a non‑existent built‑in document property.
// Use Cases: Verify a built‑in property exists before using it to avoid runtime crashes. | Log detailed exception information when a requested document property is unavailable. | Continue processing and persist the workbook even after a property‑access error.
// AI Prompts: Generate C# code with Aspose.Cells that safely retrieves a built‑in document property and provides a default value when the property is missing. | Show an alternative pattern to check for a built‑in property’s presence and handle the missing case without throwing an exception. | Explain how to implement robust exception handling around document property access while ensuring the workbook can still be saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDemo
{
    // This C# example creates a Workbook, tries to read a non‑existent built‑in document property, deliberately triggers a NullReferenceException, catches the error, displays its type and message, and finally saves the workbook as DemoWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance (lifecycle rule: create)
            Workbook workbook = new Workbook();

            try
            {
                // Attempt to access a built‑in property that does not exist.
                // The indexer returns null for unknown property names.
                DocumentProperty unknownProp = workbook.BuiltInDocumentProperties["NonExistentProperty"];

                // Deliberately cause a NullReferenceException by accessing a member of the null object.
                // This demonstrates exception handling for invalid property access.
                Console.WriteLine("Value of unknown property: " + unknownProp.Value);
            }
            catch (Exception ex)
            {
                // Handle the exception and display its type and message.
                Console.WriteLine("Exception caught while accessing a non‑existent built‑in property:");
                Console.WriteLine("Exception Type: " + ex.GetType().Name);
                Console.WriteLine("Message: " + ex.Message);
            }

            // Save the workbook to disk (lifecycle rule: save)
            workbook.Save("DemoWorkbook.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as DemoWorkbook.xlsx");
        }
    }
}
