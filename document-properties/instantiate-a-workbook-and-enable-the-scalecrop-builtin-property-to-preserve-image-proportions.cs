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
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Enable the ScaleCrop property to preserve image proportions
        properties.ScaleCrop = true;

        // Output the current value to verify
        Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

        // Save the workbook to a file
        workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
    }
}