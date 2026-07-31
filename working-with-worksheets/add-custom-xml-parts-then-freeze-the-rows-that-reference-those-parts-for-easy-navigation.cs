// Title: Add Custom XML Parts and Freeze Header Row in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, embed two custom XML parts (one with an XSD schema, one without), write each part's ID and description into cells A2‑B3, freeze the first row to keep the header visible, and save the file as CustomXmlWithFreeze.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | custom XML parts | add XML part C# | XML schema Aspose.Cells | freeze panes | header row freeze | worksheet freeze top row | .NET Excel automation | Workbook custom XML | Excel ID column | global developers | USA C# examples
// Common Searches: How to add a custom XML part with schema in Aspose.Cells C# | Freeze the top row in an Aspose.Cells worksheet after writing data | Display custom XML part IDs in Excel using Aspose.Cells .NET | C# example for embedding multiple XML parts in a workbook | Aspose.Cells freeze panes syntax
// Developer Intent: Embed custom XML parts (with optional XSD), list their IDs in a worksheet table, and keep the header row fixed for easy navigation.
// Use Cases: Store structured customer or order data as hidden XML parts while providing a visible reference table for auditors. | Create a documentation sheet that lists XML part identifiers alongside descriptions, with the header always in view. | Develop large Excel reports that include XML metadata and require stable column headings during scrolling.
// AI Prompts: Write C# code that adds multiple custom XML parts (with and without XSD) to an Aspose.Cells workbook and records each part's ID and description in the first worksheet. | Show how to freeze the first row of an Aspose.Cells worksheet after populating header cells. | Explain how to retrieve a custom XML part by its ID from a saved workbook using Aspose.Cells for .NET.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

// Demonstrates how to create a new Workbook, embed two custom XML parts (one with an XSD schema, one without), write each part's ID and description into cells A2‑B3, freeze the first row to keep the header visible, and save the file as CustomXmlWithFreeze.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Prepare XML data and optional schema for the first custom XML part
        string xmlData1 = "<Customer><Name>John Doe</Name></Customer>";
        string xmlSchema1 = "<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>" +
                            "<xs:element name='Customer'><xs:complexType>" +
                            "<xs:sequence><xs:element name='Name' type='xs:string'/></xs:sequence>" +
                            "</xs:complexType></xs:element></xs:schema>";

        // Add the first custom XML part to the workbook
        int partIndex1 = wb.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlData1), Encoding.UTF8.GetBytes(xmlSchema1));

        // Prepare XML data for a second custom XML part (no schema needed)
        string xmlData2 = "<Order><Id>12345</Id></Order>";

        // Add the second custom XML part
        int partIndex2 = wb.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlData2), null);

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Write identifiers so users can see which rows reference which custom XML parts
        ws.Cells["A1"].PutValue("Part ID");
        ws.Cells["B1"].PutValue("Description");
        ws.Cells["A2"].PutValue(wb.CustomXmlParts[partIndex1].ID);
        ws.Cells["B2"].PutValue("Customer XML");
        ws.Cells["A3"].PutValue(wb.CustomXmlParts[partIndex2].ID);
        ws.Cells["B3"].PutValue("Order XML");

        // Freeze the header row (row 1) to keep it visible while scrolling
        ws.FreezePanes(1, 0, 1, 0); // Freeze first row, no columns

        // Save the workbook
        wb.Save("CustomXmlWithFreeze.xlsx");
    }
}
