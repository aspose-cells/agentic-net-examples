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
        DocumentPropertyCollection builtInProps = workbook.Worksheets.BuiltInDocumentProperties;

        // Retrieve the total number of built‑in properties
        int builtInCount = builtInProps.Count;

        // Log the count
        Console.WriteLine($"Total built‑in document properties: {builtInCount}");

        // (Optional) Retrieve and log the count of custom properties as well
        DocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;
        Console.WriteLine($"Total custom document properties: {customProps.Count}");

        // Save the workbook (required to persist any changes)
        workbook.Save("DocumentPropertiesCount.xlsx");
    }
}