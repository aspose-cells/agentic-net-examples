using System;
using Aspose.Cells;
using Aspose.Cells.Properties; // Required for DocumentProperty

class Program
{
    static void Main()
    {
        // Instantiate a new workbook (creation rule)
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to retrieve a built‑in property that does not exist.
            // The indexer returns null when the name is not a recognized built‑in property.
            DocumentProperty prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];

            // Accessing .Value on a null reference will throw a NullReferenceException.
            // This demonstrates exception handling for an invalid property access.
            prop.Value = "Some value";
        }
        catch (Exception ex)
        {
            // Catch and display the exception details.
            Console.WriteLine($"Exception caught: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
        }

        // Save the workbook to disk (save rule)
        workbook.Save("DemoWorkbook.xlsx");
    }
}