// Title: Add a Custom XML Part and Freeze the Header Row with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a custom XML part (no schema), assigns a GUID as its ID, writes the ID into cells A1‑B1 of the first worksheet, freezes the top row so the reference stays visible while scrolling, and saves the file as CustomXmlWithFreeze.xlsx.
// Keywords: Aspose.Cells custom XML part C# | add custom XML part Aspose.Cells | freeze first row Aspose.Cells | worksheet freeze panes .NET | store XML metadata in workbook | GUID custom XML part ID | Aspose.Cells example .NET | C# Excel custom XML part
// Common Searches: how to add a custom XML part in Aspose.Cells .NET | freeze header row after writing data with Aspose.Cells | write custom XML part ID to a cell using Aspose.Cells | Aspose.Cells C# freeze panes example | retrieve custom XML part GUID in Aspose.Cells workbook
// Developer Intent: Insert a custom XML part, display its GUID in a worksheet cell, and keep that row fixed for constant visibility.
// Use Cases: Embed metadata as a custom XML part and show its ID in a frozen top row for quick reference. | Generate reports where each sheet lists its linked XML part ID in a non‑scrollable header. | Create an audit trail by recording XML part IDs in a frozen row, ensuring they remain visible during data review.
// AI Prompts: Write C# code with Aspose.Cells that adds multiple custom XML parts, lists each part's GUID in separate rows, and freezes all header rows containing the IDs. | Show how to open a saved workbook and retrieve a custom XML part by its GUID using Aspose.Cells for .NET. | Provide an example that updates the content of a custom XML part, refreshes the displayed ID, and maintains the frozen header row.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Creates a new workbook, inserts a custom XML part (no schema), assigns a GUID as its ID, writes the ID into cells A1‑B1 of the first worksheet, freezes the top row so the reference stays visible while scrolling, and saves the file as CustomXmlWithFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare custom XML data (no schema in this example)
        string xmlContent = "<MyData><Item>Sample Value</Item></MyData>";
        byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);
        byte[] schemaBytes = null; // schema is optional

        // Add the custom XML part to the workbook
        int partIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

        // Assign a unique ID to the part for easy identification
        workbook.CustomXmlParts[partIndex].ID = Guid.NewGuid().ToString();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Write the custom XML part ID into the first row (as a reference)
        sheet.Cells["A1"].PutValue("Custom XML Part ID:");
        sheet.Cells["B1"].PutValue(workbook.CustomXmlParts[partIndex].ID);

        // Freeze the first row so the reference stays visible while scrolling
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        sheet.FreezePanes(1, 0, 1, 0); // Freeze row 1 (index 1) with 1 frozen row

        // Save the workbook
        workbook.Save("CustomXmlWithFreeze.xlsx");
    }
}
