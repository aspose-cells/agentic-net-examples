// Title: Aspose.Cells for .NET – Delete a ListObject (Excel Table) Comment
// Description: Creates a workbook, adds sample data, defines a ListObject covering A1:B3, assigns a temporary comment, then removes the comment by clearing the ListObject.Comment property and saves the file as TableCommentRemoved.xlsx.
// Keywords: Aspose.Cells delete table comment | clear ListObject comment C# | remove Excel table metadata Aspose.Cells | Aspose.Cells ListObject.Comment | C# Aspose.Cells table comment removal
// Common Searches: how to clear a comment from an Aspose.Cells table in C# | Aspose.Cells remove ListObject comment example | delete table comment .NET Aspose.Cells
// Developer Intent: The developer needs to erase a comment attached to a ListObject (Excel table) in a workbook using Aspose.Cells for .NET.
// Use Cases: Strip temporary notes from generated tables before sharing a report. | Sanitize workbook metadata during an automated cleanup pipeline. | Prepare workbooks for archiving by removing all table comments to meet compliance standards.
// AI Prompts: Show how to delete a ListObject comment in Aspose.Cells without altering other table settings. | Provide a C# loop that clears comments from every table in a worksheet using Aspose.Cells. | Explain the effect of setting ListObject.Comment to null versus string.Empty in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds sample data, defines a ListObject covering A1:B3, assigns a temporary comment, then removes the comment by clearing the ListObject.Comment property and saves the file as TableCommentRemoved.xlsx.
class DeleteTableComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the table
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue(200);

        // Create a ListObject (table) that covers the data range A1:B3
        int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Assign a comment to the table (metadata before review)
        table.Comment = "Table created for initial analysis";

        // Delete the comment after final review by clearing the property
        table.Comment = string.Empty; // or null

        // Save the workbook with the cleaned‑up table metadata
        workbook.Save("TableCommentRemoved.xlsx", SaveFormat.Xlsx);
    }
}
