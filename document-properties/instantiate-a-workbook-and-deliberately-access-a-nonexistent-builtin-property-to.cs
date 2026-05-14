using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook instance (uses the provided constructor rule)
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a built‑in property that does not exist.
            // "NonExistentProperty" is not a recognized built‑in name, so the indexer returns null.
            // Accessing .Value on a null reference will throw a NullReferenceException.
            DocumentProperty prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];
            prop.Value = "Test"; // This line triggers the exception.
        }
        catch (Exception ex)
        {
            // Handle the exception and display details.
            Console.WriteLine("Exception caught while accessing a non‑existent built‑in property:");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Exception Type: {ex.GetType().Name}");
        }

        // Save the workbook (uses the provided Save rule)
        workbook.Save("DemoWorkbook.xlsx");
    }
}