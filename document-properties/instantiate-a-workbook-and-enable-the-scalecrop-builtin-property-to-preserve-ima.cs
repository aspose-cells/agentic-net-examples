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

            // Enable the ScaleCrop built‑in property to preserve image proportions
            workbook.BuiltInDocumentProperties.ScaleCrop = true;

            // Optional: display the current value to verify
            Console.WriteLine("ScaleCrop property is set to: " + workbook.BuiltInDocumentProperties.ScaleCrop);

            // Save the workbook to a file (XLSX format)
            workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}