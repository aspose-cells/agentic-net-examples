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

        // Disable thumbnail cropping by setting ScaleCrop to false
        properties.ScaleCrop = false;

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}