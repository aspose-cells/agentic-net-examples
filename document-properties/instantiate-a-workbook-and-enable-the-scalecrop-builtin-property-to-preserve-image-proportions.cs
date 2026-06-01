using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ScaleCropDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Enable ScaleCrop to preserve image proportions in the thumbnail
        properties.ScaleCrop = true;

        // Output the current value to verify
        Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

        // Save the workbook to a file (XLSX format)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}