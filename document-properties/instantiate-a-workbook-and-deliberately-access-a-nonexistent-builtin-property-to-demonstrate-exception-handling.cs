using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        try
        {
            // Attempt to access a built‑in property that does not exist.
            // The indexer returns null for unknown property names.
            DocumentProperty prop = workbook.BuiltInDocumentProperties["NonExistentProperty"];

            // Force an exception by accessing a member of the null reference.
            Console.WriteLine("Value: " + prop.Value);
        }
        catch (NullReferenceException ex)
        {
            // Expected when the property name is not recognized.
            Console.WriteLine("Caught NullReferenceException: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Fallback for any other unexpected errors.
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
        finally
        {
            // Release resources.
            workbook.Dispose();
        }
    }
}