// Title: Export Custom Document Properties to a "Metadata" Sheet with Aspose.Cells for .NET (C#)
// Description: Loads an existing Excel file, reads its custom document properties, creates a new worksheet named "Metadata", writes a header row, lists each property name and value in two columns, and saves the result to a new workbook.
// Keywords: Aspose.Cells read custom properties C# | export custom document properties Excel | create metadata worksheet Aspose.Cells | list workbook custom properties .NET | write property name and value to sheet
// Common Searches: Aspose.Cells export custom properties to new sheet | C# add Metadata worksheet with custom document properties | how to list Excel custom properties using Aspose.Cells | save custom document properties as table in Excel
// Developer Intent: Read all custom document properties from a workbook and write them into a newly added "Metadata" worksheet.
// Use Cases: Generate a quick‑reference sheet for auditing workbook metadata. | Create a portable copy that bundles data and its custom properties for downstream processing. | Automate documentation of workbook settings before distribution to stakeholders.
// AI Prompts: Show C# code with Aspose.Cells that exports custom document properties to a CSV file. | Add robust error handling for workbooks that contain no custom properties. | Demonstrate how to format the Metadata sheet (bold headers, auto‑fit columns, borders) using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataUtility
{
    // Loads an existing Excel file, reads its custom document properties, creates a new worksheet named "Metadata", writes a header row, lists each property name and value in two columns, and saves the result to a new workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source and destination workbooks
            string sourcePath = "input.xlsx";
            string destinationPath = "output.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Retrieve the collection of custom document properties
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Add a new worksheet named "Metadata"
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet metadataSheet = workbook.Worksheets[sheetIndex];
            metadataSheet.Name = "Metadata";

            // Write header titles
            metadataSheet.Cells["A1"].PutValue("Property Name");
            metadataSheet.Cells["B1"].PutValue("Value");

            // Populate the worksheet with custom property names and values
            int row = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (DocumentProperty prop in customProps)
            {
                metadataSheet.Cells[row, 0].PutValue(prop.Name);
                metadataSheet.Cells[row, 1].PutValue(prop.Value?.ToString() ?? string.Empty);
                row++;
            }

            // Save the modified workbook
            workbook.Save(destinationPath);
        }
    }
}
