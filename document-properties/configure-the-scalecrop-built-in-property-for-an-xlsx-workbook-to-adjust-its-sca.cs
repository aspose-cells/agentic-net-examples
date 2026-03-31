using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsScaleCropDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access built‑in document properties
            BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

            // Enable ScaleCrop to adjust the thumbnail scaling behavior
            properties.ScaleCrop = true;

            // Verify the property value
            Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

            // Save the workbook as an XLSX file
            workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}