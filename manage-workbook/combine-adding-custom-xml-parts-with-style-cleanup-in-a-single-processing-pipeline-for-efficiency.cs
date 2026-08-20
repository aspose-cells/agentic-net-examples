// Title: Add a Custom XML Part and Remove Unused Styles with Aspose.Cells for .NET
// Description: Demonstrates how to embed a custom XML part (with optional schema and GUID) into a workbook, delete rows that leave orphaned styles, and then purge those unused styles in a single processing pipeline before saving the file.
// Keywords: Aspose.Cells custom XML part | RemoveUnusedStyles | C# workbook manipulation | embed XML schema in Excel | style cleanup Aspose.Cells | GUID for custom XML part | batch workbook operations
// Common Searches: Aspose.Cells add custom XML part C# | how to delete unused styles after row removal | combine XML part insertion with style cleanup | assign GUID to custom XML in Aspose.Cells | batch processing workbook Aspose.Cells .NET
// Developer Intent: Insert a custom XML segment into an Excel file and immediately eliminate any style objects that became redundant after row deletions, all within one workflow.
// Use Cases: Create a report, embed metadata as XML, delete temporary rows, and shrink the file by removing stray styles. | Prepare an archival workbook with schema‑validated XML data while ensuring only active styles remain to reduce size. | Build a template that programmatically adds exchange‑ready XML and cleans up styling after dynamic row operations.
// AI Prompts: Generate C# code using Aspose.Cells to add a custom XML part with a schema, assign a unique GUID, delete specific rows, and call RemoveUnusedStyles before saving. | Show an example that batches custom XML insertion and style cleanup in Aspose.Cells, handling byte arrays and GUID assignment. | Explain the most efficient way to perform multiple workbook modifications—XML part addition, row deletion, and unused style removal—in a single Aspose.Cells pipeline.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates how to embed a custom XML part (with optional schema and GUID) into a workbook, delete rows that leave orphaned styles, and then purge those unused styles in a single processing pipeline before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add sample data with varying styles to generate multiple style objects
        Worksheet ws = wb.Worksheets[0];
        for (int i = 0; i < 5; i++)
        {
            Cell cell = ws.Cells[i, 0];
            cell.PutValue($"Item {i + 1}");

            // Create a distinct style for each cell
            Style style = wb.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10 + i;
            style.Font.IsBold = (i % 2 == 0);
            cell.SetStyle(style);
        }

        // Delete some rows to leave behind unused styles
        ws.Cells.DeleteRows(3, 2);

        // Prepare custom XML data and optional schema
        string xmlData = "<MyData xmlns=\"http://example.com\"><Item>Value</Item></MyData>";
        string xmlSchema = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'></xs:schema>";
        byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
        byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

        // Add the custom XML part to the workbook
        int xmlIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);
        // Optionally assign a unique ID to the part
        wb.CustomXmlParts[xmlIndex].ID = Guid.NewGuid().ToString();

        // Remove all styles that are no longer used after row deletion
        wb.RemoveUnusedStyles();

        // Save the workbook with both custom XML and cleaned styles
        wb.Save("CombinedOutput.xlsx");
    }
}
