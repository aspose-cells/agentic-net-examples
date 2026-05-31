using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will become the table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Add a ListObject (Excel table) covering the data range A1:B3
        int tableIdx = sheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = sheet.ListObjects[tableIdx];
        table.DisplayName = "Employees";

        // Attach a comment to the table that contains a SharePoint hyperlink (as plain text)
        table.Comment = "Reference document: https://sharepoint.company.com/sites/docs/EmployeeGuide.docx";

        // Add a clickable hyperlink to the first cell of the table (A1) pointing to the same SharePoint document
        int hyperlinkIdx = sheet.Hyperlinks.Add("A1", 1, 1, "https://sharepoint.company.com/sites/docs/EmployeeGuide.docx");
        Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIdx];
        hyperlink.TextToDisplay = "Employee Guide";
        hyperlink.ScreenTip = "Open SharePoint document";

        // Save the workbook
        workbook.Save("TableWithCommentAndHyperlink.xlsx", SaveFormat.Xlsx);
    }
}