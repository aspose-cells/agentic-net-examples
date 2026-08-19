// Title: Aspose.Cells C# – Export Custom Document Properties to a Metadata Worksheet
// Description: Load an Excel file with Aspose.Cells for .NET, read its custom document properties, create a new sheet named "Metadata", list each property name and value, and save the updated workbook.
// Keywords: Aspose.Cells read custom properties | C# export Excel metadata | custom document properties worksheet | Aspose.Cells create metadata sheet | write custom properties to Excel
// Common Searches: how to export custom document properties using Aspose.Cells | Aspose.Cells .NET create metadata tab from workbook properties | C# list custom Excel properties in a new sheet | Aspose.Cells read and write custom document properties
// Developer Intent: Read a workbook’s custom document properties and write them to a newly added "Metadata" worksheet.
// Use Cases: Generate a quick reference sheet of all custom properties for audit trails. | Provide downstream processes with a standardized metadata tab for configuration values. | Export property data before sharing the file so recipients can view custom settings without opening the properties dialog.
// AI Prompts: Generate C# code with Aspose.Cells that reads every custom document property from a workbook and writes name/value pairs to a new worksheet called "Metadata". | Show how to extend the utility to include built‑in document properties alongside custom ones in the same metadata sheet. | Explain how to format non‑string property values (dates, numbers, booleans) correctly when writing them to cells with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataUtility
{
    // Load an Excel file with Aspose.Cells for .NET, read its custom document properties, create a new sheet named "Metadata", list each property name and value, and save the updated workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains custom document properties
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle: use provided constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Access the collection of custom document properties
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Add a new worksheet to hold the metadata
            int metaSheetIndex = workbook.Worksheets.Add();
            Worksheet metaSheet = workbook.Worksheets[metaSheetIndex];
            metaSheet.Name = "Metadata";

            // Write header row
            metaSheet.Cells["A1"].PutValue("Property Name");
            metaSheet.Cells["B1"].PutValue("Value");

            // Populate the worksheet with each custom property (Name in column A, Value in column B)
            for (int i = 0; i < customProps.Count; i++)
            {
                DocumentProperty prop = customProps[i];
                // Row index in Cells is zero‑based; add 1 for header offset
                int row = i + 1;
                metaSheet.Cells[row, 0].PutValue(prop.Name);
                metaSheet.Cells[row, 1].PutValue(prop.Value);
            }

            // Save the modified workbook (lifecycle: use provided Save method)
            workbook.Save("output.xlsx");
        }
    }
}
