// Title: Set ScaleCrop Built‑In Document Property in Aspose.Cells C# Workbook
// Description: Creates a new Aspose.Cells workbook, enables the ScaleCrop property to keep image aspect ratios, prints the value, and saves the file as XLSX.
// Keywords: Aspose.Cells | ScaleCrop | C# | BuiltInDocumentProperties | preserve image aspect ratio | Workbook.Save | Excel | document property | image scaling | .NET
// Common Searches: Aspose.Cells set ScaleCrop property C# | How to preserve image proportions in Aspose.Cells workbook | Enable ScaleCrop built‑in document property Aspose.Cells .NET | ScaleCrop Aspose.Cells example | C# Aspose.Cells image scaling
// Developer Intent: Enable the ScaleCrop built‑in document property so images retain their original aspect ratio when the workbook is saved.
// Use Cases: Maintain correct image aspect ratios when generating reports with embedded pictures. | Standardize image scaling across Excel viewers for programmatically created workbooks. | Apply a global image scaling setting before exporting a workbook that contains multiple graphics.
// AI Prompts: Generate C# code using Aspose.Cells to create a workbook, set ScaleCrop = true, insert an image, and save as .xlsx. | Explain the effect of the ScaleCrop property on image rendering in an Aspose.Cells workbook and show how to toggle it. | Provide a step‑by‑step example that reads the ScaleCrop value, changes it, and verifies the change at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsScaleCropDemo
{
    // Creates a new Aspose.Cells workbook, enables the ScaleCrop property to keep image aspect ratios, prints the value, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the built‑in document properties collection
            BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

            // Enable ScaleCrop to preserve image proportions
            properties.ScaleCrop = true;

            // Optional: display the current value to verify
            Console.WriteLine("ScaleCrop property value: " + properties.ScaleCrop);

            // Save the workbook (uses the provided Save method rule)
            workbook.Save("ScaleCropDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
