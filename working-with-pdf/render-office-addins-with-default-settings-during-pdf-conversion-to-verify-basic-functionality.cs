using System;
using Aspose.Cells;
using Aspose.Cells.WebExtensions;
using Aspose.Cells.Rendering;

namespace OfficeAddInPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the WebExtensions collection of the first worksheet
            WebExtensionCollection webExtensions = workbook.Worksheets.WebExtensions;

            // Add a new WebExtension (Office Add‑In) with default settings
            int extIndex = webExtensions.Add();                     // creates a new WebExtension
            WebExtension webExt = webExtensions[extIndex];

            // Set minimal required reference information (default values)
            webExt.Reference.Id = "sampleAddIn";
            webExt.Reference.StoreName = "SampleStore";

            // Optionally add a property to the add‑in (not required for default behavior)
            webExt.Properties.Add("exampleProperty", "exampleValue");

            // Save the workbook as PDF using default PdfSaveOptions (lifecycle: save)
            PdfSaveOptions pdfOptions = new PdfSaveOptions(); // default settings
            workbook.Save("OfficeAddInDemo.pdf", pdfOptions);

            Console.WriteLine("PDF generated with Office Add‑In using default settings.");
        }
    }
}