// Title: Disable LinksUpToDate Built‑in Property in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Load an existing .xlsx file using Aspose.Cells, set the built‑in document property LinksUpToDate to false to stop Excel from checking external links, and save the updated workbook.
// Keywords: Aspose.Cells LinksUpToDate | disable link checking Excel .NET | set built‑in document property false C# | prevent external link updates Aspose.Cells | Excel workbook properties C# | Aspose.Cells document properties example | C# LinksUpToDate false
// Common Searches: How to turn off LinksUpToDate in Aspose.Cells | Aspose.Cells C# disable link verification | Set LinksUpToDate false programmatically | Prevent Excel external links from updating with Aspose.Cells | Aspose.Cells built‑in document properties tutorial
// Developer Intent: Programmatically set the LinksUpToDate built‑in property to false so the workbook does not perform link validation.
// Use Cases: Distribute templates that contain external references without prompting users to refresh links. | Improve performance when opening large workbooks that reference data sources not needed during processing. | Create offline reports where link updates are irrelevant and should be suppressed.
// AI Prompts: Write a C# example that loads an .xlsx file with Aspose.Cells, disables the LinksUpToDate property, and saves the result. | Explain why setting LinksUpToDate to false stops Excel from checking external links and how to implement it using Aspose.Cells. | Provide step‑by‑step code to prevent link verification in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Load an existing .xlsx file using Aspose.Cells, set the built‑in document property LinksUpToDate to false to stop Excel from checking external links, and save the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook file
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Disable the LinksUpToDate built‑in property to prevent link checks
            workbook.BuiltInDocumentProperties.LinksUpToDate = false;

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded from '{inputPath}', LinksUpToDate set to false, and saved as '{outputPath}'.");
        }
    }
}
