// Title: Add a Link‑to‑Content Custom Property "SourceUrl" from Cell A1 using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, writes an external URL into cell A1, adds a custom document property named "SourceUrl" that links to that cell, updates the linked value, displays it, and saves the file as WorkbookWithLinkedProperty.xlsx.
// Keywords: Aspose.Cells link to content property | custom document property from cell | C# Aspose.Cells linked property | UpdateLinkedPropertyValue example | save workbook with custom property
// Common Searches: how to add a linked custom property in Aspose.Cells | Aspose.Cells set custom document property from cell value | update linked property after cell change Aspose.Cells | C# create link‑to‑content property in Excel file
// Developer Intent: Create a workbook and attach a link‑to‑content custom document property called "SourceUrl" that reflects the URL stored in cell A1.
// Use Cases: Track the source URL of data used to generate the spreadsheet for audit trails. | Keep a property automatically synchronized when the referenced cell is edited. | Expose external references in file metadata for downstream processing without opening the workbook.
// AI Prompts: Write C# code with Aspose.Cells that adds a link‑to‑content custom property referencing a cell containing a URL and saves the workbook. | Explain why UpdateLinkedPropertyValue is required and when to call it after modifying the source cell. | Show how to read a linked custom document property from an existing workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a new workbook, writes an external URL into cell A1, adds a custom document property named "SourceUrl" that links to that cell, updates the linked value, displays it, and saves the file as WorkbookWithLinkedProperty.xlsx.
class AddLinkToContentProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put the external URL into a cell (this will be the source of the linked property)
        string externalUrl = "https://www.example.com/data";
        sheet.Cells["A1"].PutValue(externalUrl);

        // Add a custom document property that links to the content of cell A1
        // The property name is "SourceUrl" and the source is the cell reference "A1"
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;
        customProps.AddLinkToContent("SourceUrl", "A1");

        // Update the linked property value so it reflects the current cell content
        customProps.UpdateLinkedPropertyValue();

        // Optionally, display the linked property's value to verify
        Console.WriteLine("SourceUrl property value: " + customProps["SourceUrl"].Value);

        // Save the workbook
        workbook.Save("WorkbookWithLinkedProperty.xlsx");
    }
}
