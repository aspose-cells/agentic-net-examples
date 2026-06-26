using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        try
        {
            // Try to get a built‑in document property that does not exist.
            // The indexer returns null for unknown property names.
            var unknownProperty = workbook.BuiltInDocumentProperties["NonExistentProperty"];

            // Accessing Value on a null reference will throw a NullReferenceException.
            Console.WriteLine("Property Value: " + unknownProperty.Value);
        }
        catch (Exception ex)
        {
            // Handle the exception and display its details.
            Console.WriteLine("Exception caught:");
            Console.WriteLine("Type: " + ex.GetType().Name);
            Console.WriteLine("Message: " + ex.Message);
        }

        // Save the workbook to complete the normal lifecycle.
        workbook.Save("DemoWorkbook.xlsx");
    }
}