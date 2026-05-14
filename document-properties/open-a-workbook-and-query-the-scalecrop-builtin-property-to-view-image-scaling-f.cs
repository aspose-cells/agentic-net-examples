using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Query the ScaleCrop property which indicates the thumbnail display mode
        bool scaleCrop = properties.ScaleCrop;

        // Output the current value of ScaleCrop
        Console.WriteLine("ScaleCrop property value: " + scaleCrop);
    }
}