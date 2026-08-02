using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        DocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Access the custom document properties collection
        DocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;

        // Retrieve the total number of properties in each collection
        int builtInCount = builtInProps.Count;
        int customCount = customProps.Count;

        // Output the counts to the console (log)
        Console.WriteLine($"Built‑in document properties count: {builtInCount}");
        Console.WriteLine($"Custom document properties count: {customCount}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("DocumentPropertiesCount.xlsx");
    }
}