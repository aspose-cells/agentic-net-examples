// Title: Add a Link‑to‑Content Custom Property “SourceUrl” in Aspose.Cells (C#)
// Description: Creates a new Workbook, writes an external URL into cell A1, adds a custom document property named SourceUrl that links to that cell using AddLinkToContent, synchronizes the linked value, and saves the file as WorkbookWithLinkedUrl.xlsx.
// Keywords: Aspose.Cells AddLinkToContent | C# custom document property linked to cell | store URL as workbook property | UpdateLinkedPropertyValue example | Aspose.Cells custom properties tutorial
// Common Searches: how to add a link‑to‑content custom property in Aspose.Cells .NET | Aspose.Cells C# store cell value as custom document property | synchronize linked custom properties after cell change Aspose.Cells | AddLinkToContent usage sample | retrieve linked custom property value Aspose.Cells
// Developer Intent: Generate a workbook, place an external URL in a cell, and create a linked custom document property called SourceUrl that reflects the cell’s content.
// Use Cases: Reference an external data source URL from anywhere in the workbook without hard‑coding it in formulas. | Keep the property value automatically updated when the source cell is edited. | Expose the URL to downstream systems that read document properties instead of parsing worksheet cells.
// AI Prompts: Write C# code with Aspose.Cells to add a link‑to‑content custom property that points to a cell containing an external URL. | Explain the purpose of UpdateLinkedPropertyValue and when it should be invoked. | Show how to read the value of a linked custom document property from an existing workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a new Workbook, writes an external URL into cell A1, adds a custom document property named SourceUrl that links to that cell using AddLinkToContent, synchronizes the linked value, and saves the file as WorkbookWithLinkedUrl.xlsx.
class AddLinkToContentCustomProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the external URL to be stored
        string externalUrl = "https://www.example.com/data";

        // Place the URL in a cell (A1) – this cell will be the source for the linked property
        sheet.Cells["A1"].PutValue(externalUrl);

        // Add a custom document property that links to the content of cell A1
        // The property name is "SourceUrl" and the source is the cell reference "A1"
        workbook.CustomDocumentProperties.AddLinkToContent("SourceUrl", "A1");

        // (Optional) Update linked properties to ensure the value is synchronized
        workbook.CustomDocumentProperties.UpdateLinkedPropertyValue();

        // Save the workbook to a file
        workbook.Save("WorkbookWithLinkedUrl.xlsx");
    }
}
