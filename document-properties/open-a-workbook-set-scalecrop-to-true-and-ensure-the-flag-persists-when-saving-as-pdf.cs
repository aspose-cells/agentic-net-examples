// Title: C# – Set Workbook ScaleCrop Flag and Keep It When Saving as PDF with Aspose.Cells
// Description: Shows how to turn on the BuiltInDocumentProperties.ScaleCrop attribute for an Aspose.Cells workbook, export the workbook to PDF, and confirm that the flag remains set in the generated file.
// Keywords: Aspose.Cells | C# | .NET | ScaleCrop | BuiltInDocumentProperties | PDF export | document properties | thumbnail scaling | Excel to PDF | preserve settings
// Common Searches: Aspose.Cells set ScaleCrop | ScaleCrop flag PDF export | keep ScaleCrop after saving PDF | thumbnail scaling property Aspose.Cells | how to enable ScaleCrop in C#
// Developer Intent: Enable the ScaleCrop attribute on a workbook’s built‑in properties and ensure the setting survives the conversion to PDF.
// Use Cases: Generate PDFs with correctly scaled thumbnails for single or batch Excel files. | Modify existing workbooks to control thumbnail appearance before distribution. | Automate conversion pipelines where document‑property consistency is required.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, sets BuiltInDocumentProperties.ScaleCrop to true, saves as PDF, and prints the property value. | Explain the effect of the ScaleCrop property on PDF thumbnails and how to test its persistence. | Create a reusable function that accepts input and output paths, applies ScaleCrop, and returns the final property state.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Shows how to turn on the BuiltInDocumentProperties.ScaleCrop attribute for an Aspose.Cells workbook, export the workbook to PDF, and confirm that the flag remains set in the generated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path) or create a new one
        // Workbook workbook = new Workbook("input.xlsx");
        Workbook workbook = new Workbook();

        // Set the ScaleCrop property to true so the thumbnail is displayed in scaled mode
        workbook.BuiltInDocumentProperties.ScaleCrop = true;

        // Save the workbook as PDF; the ScaleCrop flag is persisted in the document properties
        workbook.Save("output.pdf", SaveFormat.Pdf);

        // Output the current value to confirm it was set
        Console.WriteLine("ScaleCrop property value: " + workbook.BuiltInDocumentProperties.ScaleCrop);
    }
}
