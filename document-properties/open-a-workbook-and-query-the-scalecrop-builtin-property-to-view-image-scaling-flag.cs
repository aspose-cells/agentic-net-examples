using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Query the ScaleCrop property (true indicates the thumbnail is scaled)
        bool scaleCrop = properties.ScaleCrop;

        // Output the property value to the console
        Console.WriteLine("ScaleCrop property value: " + scaleCrop);

        // Save the workbook (no modifications made) to demonstrate proper lifecycle handling
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}