using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        DocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Access the custom document properties collection
        DocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;

        // Retrieve counts
        int builtInCount = builtInProps.Count;
        int customCount = customProps.Count;
        int totalCount = builtInCount + customCount;

        // Log the counts
        Console.WriteLine($"Built‑in properties count: {builtInCount}");
        Console.WriteLine($"Custom properties count: {customCount}");
        Console.WriteLine($"Total document properties count: {totalCount}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("DocumentPropertiesCount.xlsx");
    }
}