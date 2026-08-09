// Title: Aspose.Cells C# – Enable ScaleCrop Built‑in Property to Keep Workbook Thumbnail Proportions
// Description: A concise C# sample that creates an Aspose.Cells Workbook, accesses its BuiltInDocumentPropertyCollection, sets ScaleCrop to true so the generated thumbnail retains its original aspect ratio, prints the setting for verification, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells ScaleCrop | C# built‑in document properties | preserve thumbnail aspect ratio | ScaleCrop true | Aspose.Cells workbook thumbnail | document property ScaleCrop .NET | Excel thumbnail scaling | global
// Common Searches: Aspose.Cells set ScaleCrop property C# | How to keep Excel thumbnail proportions with Aspose | ScaleCrop built‑in document property example | C# preserve workbook thumbnail aspect ratio | Aspose.Cells thumbnail image scaling
// Developer Intent: Activate the ScaleCrop built‑in document property so the workbook’s thumbnail image remains proportional.
// Use Cases: Generate reports programmatically and ensure the thumbnail image is not distorted. | Create a template workbook where the thumbnail must match the original picture’s aspect ratio. | Validate the ScaleCrop setting in automated CI pipelines before publishing Excel files.
// AI Prompts: Show a C# snippet that toggles the ScaleCrop property on an existing Aspose.Cells workbook. | Explain how the ScaleCrop setting influences Excel thumbnail rendering and how to confirm its value. | Provide an example that reads the current ScaleCrop flag, changes it, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// A concise C# sample that creates an Aspose.Cells Workbook, accesses its BuiltInDocumentPropertyCollection, sets ScaleCrop to true so the generated thumbnail retains its original aspect ratio, prints the setting for verification, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Enable ScaleCrop to preserve image proportions in the thumbnail
        properties.ScaleCrop = true;

        // Output the current setting (optional verification)
        Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

        // Save the workbook to a file
        workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
    }
}
