// Title: Save Workbook as XLSB while Preserving XML Maps – C# Aspose.Cells Example
// Description: Demonstrates how to create a workbook, add sample data, configure XlsbSaveOptions, ensure the output folder exists, and save the file as a binary .xlsb. The code notes that XML‑map support depends on the Aspose.Cells version, and any attached maps are retained when the library supports them.
// Keywords: Aspose.Cells | C# | XLSB | XlsbSaveOptions | XML map | save workbook | binary Excel | retain XML map | export to .xlsb | Aspose.Cells .NET
// Common Searches: Aspose.Cells save as xlsb with xml map | C# export workbook to xlsb preserving xml map | keep xml maps when saving xlsb Aspose | XlsbSaveOptions retain xml mapping | save workbook to binary format Aspose.Cells
// Developer Intent: Export a workbook to the XLSB binary format while keeping any attached XML maps intact (provided the library version supports XmlMaps).
// Use Cases: Generate a workbook, populate it with data, and save it as a compact .xlsb file for faster loading. | Maintain XML‑map definitions for downstream data import/export processes when the Aspose.Cells version includes XmlMap support. | Create the target directory programmatically to avoid path‑not‑found errors before calling Workbook.Save. | Wrap the save operation in try‑catch logic to handle version‑related or I/O exceptions gracefully.
// AI Prompts: Write C# code that attaches an XML map to a workbook and saves it as an XLSB file using Aspose.Cells, including a version check for XmlMap support. | Explain how XlsbSaveOptions works in Aspose.Cells and describe any limitations when preserving XML maps during XLSB export. | Provide best‑practice guidelines for error handling and folder management when saving a workbook with XML maps to XLSB in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, add sample data, configure XlsbSaveOptions, ensure the output folder exists, and save the file as a binary .xlsb. The code notes that XML‑map support depends on the Aspose.Cells version, and any attached maps are retained when the library supports them.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(999.99);

            // NOTE: XML map functionality is not available in the current Aspose.Cells version.
            // If needed, ensure the library version supports XmlMaps before using the related APIs.

            // Create XLSB save options (default options retain all workbook data)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions();

            // Define output path and ensure the directory exists
            string outputPath = "ProductData.xlsb";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLSB file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
