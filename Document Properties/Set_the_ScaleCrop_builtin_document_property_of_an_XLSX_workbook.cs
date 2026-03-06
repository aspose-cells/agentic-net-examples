using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsScaleCropDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access built‑in document properties
            BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

            // Set the ScaleCrop property (true = thumbnail will be scaled and cropped)
            properties.ScaleCrop = true;

            // Optional: display the current value to verify
            Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

            // Save the workbook as XLSX
            workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}