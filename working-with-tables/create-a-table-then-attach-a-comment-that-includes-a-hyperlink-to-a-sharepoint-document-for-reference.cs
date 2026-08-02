// Title: Aspose.Cells C# – Insert a SharePoint URL into an Excel Table Comment
// Description: This C# example shows how to create a workbook with Aspose.Cells, add a ListObject (Excel table) over range A1:B3, set its DisplayName, assign a comment that contains a SharePoint document URL (Excel renders the URL as a clickable link), and save the file as TableWithCommentAndHyperlink.xlsx.
// Keywords: Aspose.Cells C# | Excel table comment hyperlink | ListObject comment SharePoint | Add comment to Excel table .NET | Aspose.Cells table with URL | C# Excel comment hyperlink | SharePoint link in Excel table comment
// Common Searches: Aspose.Cells add comment with URL | C# insert SharePoint link into Excel table comment | How to set ListObject comment Aspose.Cells | Excel table comment hyperlink using Aspose.Cells | Create Excel table and attach SharePoint reference in .NET
// Developer Intent: Add a clickable SharePoint URL to an Excel table comment and save the workbook.
// Use Cases: Provide end‑users with instant access to a specification document stored on SharePoint from within the table comment. | Maintain an audit trail by linking each data table to its governing policy document hosted on SharePoint. | Enable data‑validation teams to reference external guidelines directly from the Excel file without leaving the spreadsheet.
// AI Prompts: Generate C# code using Aspose.Cells that creates a ListObject and sets its Comment property to include a SharePoint hyperlink. | Explain how the Comment property in Aspose.Cells automatically converts URLs into clickable links in Excel. | Show how to format a table comment so the displayed text is friendly while embedding a functional SharePoint URL.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentWithHyperlink
{
    // This C# example shows how to create a workbook with Aspose.Cells, add a ListObject (Excel table) over range A1:B3, set its DisplayName, assign a comment that contains a SharePoint document URL (Excel renders the URL as a clickable link), and save the file as TableWithCommentAndHyperlink.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Add a ListObject (Excel table) covering the data range A1:B3
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Attach a comment to the table that includes a SharePoint hyperlink
            // The Comment property is plain text; embedding the URL makes it clickable in Excel.
            string sharepointUrl = "https://sharepoint.company.com/sites/docs/ReferenceDoc.docx";
            table.Comment = $"Reference document: {sharepointUrl}";

            // Save the workbook
            workbook.Save("TableWithCommentAndHyperlink.xlsx", SaveFormat.Xlsx);
        }
    }
}
