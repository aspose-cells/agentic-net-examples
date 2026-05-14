using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new workbook instance

        // Access built‑in document properties and set ScaleCrop to true
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;
        properties.ScaleCrop = true;

        // Verify that the property is set (optional)
        Console.WriteLine("ScaleCrop is set to: " + properties.ScaleCrop);

        // Prepare PDF save options (default options are sufficient to retain the property)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF; the ScaleCrop flag will be persisted in the output file
        workbook.Save("output.pdf", pdfOptions);
    }
}