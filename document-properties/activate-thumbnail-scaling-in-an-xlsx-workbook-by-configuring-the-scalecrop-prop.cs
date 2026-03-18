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
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Activate thumbnail scaling by setting ScaleCrop to true
        properties.ScaleCrop = true;

        // Optional: display the current value to verify
        Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}