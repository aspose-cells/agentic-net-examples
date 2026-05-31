using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

namespace AsposeCellsScaleCropPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the built‑in document properties collection
            BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

            // Set the ScaleCrop flag to true – this controls the display mode of the document thumbnail
            properties.ScaleCrop = true;

            // Optional: verify that the property is set
            Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

            // Save the workbook as a PDF.
            // The ScaleCrop flag is stored in the document properties and will be persisted in the PDF output.
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook saved as PDF with ScaleCrop enabled.");
        }
    }
}