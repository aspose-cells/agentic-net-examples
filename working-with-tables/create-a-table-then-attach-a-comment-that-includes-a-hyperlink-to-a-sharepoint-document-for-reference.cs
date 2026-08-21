// Title: Add a SharePoint hyperlink to an Excel table comment with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new workbook, fill a range with data, insert a ListObject (Excel table), assign a plain‑text comment that contains a SharePoint document URL, optionally add a visible cell hyperlink, and save the file as an XLSX workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# example | Excel table comment hyperlink | SharePoint URL in Aspose.Cells | ListObject comment .NET | C# add hyperlink to Excel table | Aspose.Cells GitHub sample | Excel automation SharePoint link | Aspose.Cells API tutorial | C# .NET Excel automation | Office Open XML hyperlink
// Common Searches: how to set a comment with a SharePoint link on a ListObject using Aspose.Cells | Aspose.Cells C# add hyperlink inside an Excel table comment | example of Excel table comment containing a URL with Aspose.Cells | C# Aspose.Cells add SharePoint hyperlink to table comment | Aspose.Cells create table and attach reference link
// Developer Intent: Insert a ListObject into a worksheet and embed a SharePoint document URL in its comment.
// Use Cases: Automated reporting where the table comment points users to a detailed policy stored on SharePoint. | Building a reusable Excel template that includes a reference link to a corporate guideline hosted on SharePoint. | Generating workbooks that provide both a comment link and an in‑cell hyperlink for quick navigation to supporting documents.
// AI Prompts: Generate C# code with Aspose.Cells that creates an Excel table, adds a comment containing a SharePoint document URL, and saves the workbook. | Show how to attach a plain‑text comment with a SharePoint hyperlink to a ListObject using Aspose.Cells for .NET. | Provide an Aspose.Cells example that adds both a table comment with a SharePoint link and a visible cell hyperlink in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentWithHyperlink
{
    // This example shows how to create a new workbook, fill a range with data, insert a ListObject (Excel table), assign a plain‑text comment that contains a SharePoint document URL, optionally add a visible cell hyperlink, and save the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (A1:B4)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");

            // Add a ListObject (Excel table) covering the data range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a comment on the table that includes a SharePoint hyperlink
            // The comment is plain text; the hyperlink is represented as a URL string.
            table.Comment = "Reference document: https://sharepoint.example.com/sites/Docs/Reference.docx";

            // Optionally, you could also add a visible hyperlink to a cell inside the table
            // sheet.Hyperlinks.Add("B2", 1, 1, "https://sharepoint.example.com/sites/Docs/Reference.docx");

            // Save the workbook
            workbook.Save("TableWithCommentAndHyperlink.xlsx", SaveFormat.Xlsx);
        }
    }
}
