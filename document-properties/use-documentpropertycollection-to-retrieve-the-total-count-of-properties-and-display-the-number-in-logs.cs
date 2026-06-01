using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection (inherits DocumentPropertyCollection)
        DocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Add a few custom properties to demonstrate counting
        workbook.Worksheets.CustomDocumentProperties.Add("Author", "John Doe");
        workbook.Worksheets.CustomDocumentProperties.Add("Revision", 1);
        workbook.Worksheets.CustomDocumentProperties.Add("Created", DateTime.Now);

        // Retrieve the total number of properties in each collection
        int builtInCount = builtInProps.Count;
        int customCount = workbook.Worksheets.CustomDocumentProperties.Count;

        // Output the counts to the console (log)
        Console.WriteLine($"Built‑in document properties count: {builtInCount}");
        Console.WriteLine($"Custom document properties count: {customCount}");

        // Optional: save the workbook
        workbook.Save("DocumentPropertiesCount.xlsx");
    }
}