using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsScaleCropDemo
{
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

            // Optional: display the current value of the property
            Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

            // Save the workbook to verify the setting
            workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}