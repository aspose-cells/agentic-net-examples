// Title: Aspose.Cells C# Example: Render Office Add‑In (WebExtension) in PDF with Default Settings
// Description: Creates a new Workbook, adds a minimal WebExtension (Office Add‑in), attaches it to a WebExtensionShape, and saves the file to PDF using the default PdfSaveOptions. Demonstrates basic add‑in rendering during PDF conversion with Aspose.Cells.
// Keywords: Aspose.Cells | C# | PDF conversion | WebExtension | Office Add‑in | default PdfSaveOptions | Excel to PDF | render add‑in | sample code | GitHub example
// Common Searches: Aspose.Cells render Office Add‑in in PDF | C# WebExtension shape to PDF | default PdfSaveOptions for WebExtension | export Excel with add‑in to PDF Aspose | how to include Office Add‑in in PDF using Aspose.Cells
// Developer Intent: Export an Excel workbook that contains an Office Add‑in to PDF without customizing conversion options.
// Use Cases: Quick verification that Office Add‑ins appear correctly in PDF output. | Generating documentation PDFs from templates that embed WebExtensions. | Automated testing of add‑in rendering during batch PDF conversions.
// AI Prompts: Show how to add multiple WebExtension shapes and place each on a separate PDF page with Aspose.Cells. | Explain how to enable or disable Office Add‑in rendering via PdfSaveOptions while keeping default behavior as fallback. | Provide robust error handling for missing WebExtension references during PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.WebExtensions;

// Creates a new Workbook, adds a minimal WebExtension (Office Add‑in), attaches it to a WebExtensionShape, and saves the file to PDF using the default PdfSaveOptions. Demonstrates basic add‑in rendering during PDF conversion with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook (default creation)
        Workbook workbook = new Workbook();

        // Add an Office Add‑in (WebExtension) to the workbook
        WebExtensionCollection webExtensions = workbook.Worksheets.WebExtensions;
        int extIndex = webExtensions.Add();                     // add a new WebExtension
        WebExtension webExtension = webExtensions[extIndex];

        // Set minimal reference information (default values)
        webExtension.Reference.Id = "sampleAddIn";
        webExtension.Reference.StoreName = "SampleStore";

        // Create a shape that will host the WebExtension
        ShapeCollection shapes = workbook.Worksheets[0].Shapes;
        // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
        Shape shape = shapes.AddShape(MsoDrawingType.WebExtension, 0, 0, 0, 0, 400, 300);
        WebExtensionShape webExtShape = (WebExtensionShape)shape;
        webExtShape.WebExtension = webExtension;               // associate the add‑in with the shape

        // Save the workbook to PDF using default PdfSaveOptions (no custom settings)
        workbook.Save("OfficeAddIn_Default.pdf", new PdfSaveOptions());

        Console.WriteLine("PDF generated with Office Add‑in using default settings.");
    }
}
